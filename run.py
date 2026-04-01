# -*- coding: utf-8 -*-
"""
run.py(PySide6 工程GUI版：源PDF -> AI预处理 -> 拼图算法 get_best6)

按当前流程：
0) 第1个目录：选择“源文件夹”(原始 PDF)
1) 第2个目录：选择“AI目录”(把源 PDF 复制并改后缀为 .ai 到这里，JSX 导出的 PDF 也直接输出到这里)
2) 第3个目录：选择“源文件  (备份)”(每次运行自动创建同名时间子目录，备份与本次导出 PDF 同名的源文件)
3) 第4~5个目录：输出目录保持不变
4) 子进程运行 get_best6 进行拼图

✅ 仅使用 PySide6(已去掉 PyQt5)
✅ 已移除“贴二维码”相关全部代码
✅ 已去掉“导入文件夹”步骤
✅ 仅调用 get_best6
"""

import os
import sys
import time
import shutil
import traceback
import subprocess
import re

# ====== 仅使用 get_best6(必须同目录)======
import get_best6

# ---- 强制本进程输出编码，避免混码导致子进程/父进程互相坑 ----
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_NAME_P1 = "over_test_p1.pdf"
OUT_NAME_P2 = "over_test_p2.pdf"

RUN_MODE_FULL = "full"
RUN_MODE_SEPARATE = "separate"
RUN_MODE_ALGO = "algo"

RUN_MODE_TEXT = {
    RUN_MODE_FULL: "一气呵成",
    RUN_MODE_SEPARATE: "仅分离",
    RUN_MODE_ALGO: "仅拼接",
}


def is_frozen_app():
    return bool(getattr(sys, "frozen", False))

# -------------------------
# Qt imports (ONLY PySide6)
# -------------------------
from PySide6.QtCore import Qt, QThread, Signal, Slot, QSettings
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLineEdit, QPushButton, QLabel, QTextEdit, QFileDialog,
    QMessageBox, QProgressBar, QSplitter, QCheckBox,
    QFrame, QDialog
)


# -------------------------
# 工具函数
# -------------------------
def ensure_dir(path):
    if path and (not os.path.isdir(path)):
        os.makedirs(path)


def build_time_batch_name():
    return time.strftime("%Y%m%d_%H%M%S")


def make_unique_batch_name(base_dirs, base_name=None):
    seed = (base_name or build_time_batch_name()).strip() or build_time_batch_name()
    candidate = seed
    idx = 1
    while True:
        conflict = False
        for root_dir in (base_dirs or []):
            if not root_dir:
                continue
            if os.path.exists(os.path.join(root_dir, candidate)):
                conflict = True
                break
        if not conflict:
            return candidate
        candidate = "%s_%02d" % (seed, idx)
        idx += 1


def norm_path(p):
    return os.path.normcase(os.path.normpath(os.path.abspath(p)))


def is_output_like_pdf(filename_lower):
    # 排除输出文件，避免重复处理
    if filename_lower.startswith("over_test") and filename_lower.endswith(".pdf"):
        return True
    return False


def iter_input_pdf_files(folder, recursive=True):
    """遍历源PDF(排除 over_test*.pdf)"""
    if not os.path.isdir(folder):
        return

    if recursive:
        for root, _, files in os.walk(folder):
            for fn in files:
                lf = fn.lower()
                if not lf.endswith(".pdf"):
                    continue
                if is_output_like_pdf(lf):
                    continue
                p = os.path.join(root, fn)
                if os.path.isfile(p):
                    yield p
    else:
        for fn in os.listdir(folder):
            lf = fn.lower()
            if not lf.endswith(".pdf"):
                continue
            if is_output_like_pdf(lf):
                continue
            p = os.path.join(folder, fn)
            if os.path.isfile(p):
                yield p


def list_ai_files(folder):
    """列出 folder 下的 .ai(递归)"""
    if not os.path.isdir(folder):
        return []
    out = []
    for root, _, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith(".ai"):
                p = os.path.join(root, fn)
                if os.path.isfile(p):
                    out.append(p)
    out.sort()
    return out


def snapshot_pdf_mtimes(folder):
    """
    Snapshot PDF mtimes in a folder (top-level only).
    Returns: {norm_path: (abs_path, mtime, size)}
    """
    out = {}
    if not folder or (not os.path.isdir(folder)):
        return out
    try:
        names = os.listdir(folder)
    except Exception:
        return out

    for fn in names:
        if not fn.lower().endswith(".pdf"):
            continue
        p = os.path.abspath(os.path.join(folder, fn))
        if not os.path.isfile(p):
            continue
        try:
            out[norm_path(p)] = (p, os.path.getmtime(p), os.path.getsize(p))
        except Exception:
            pass
    return out


def snapshot_pdf_mtimes_multi(folders):
    out = {}
    seen = set()
    for folder in folders or []:
        if not folder:
            continue
        nk = norm_path(folder)
        if nk in seen:
            continue
        seen.add(nk)
        out.update(snapshot_pdf_mtimes(folder))
    return out


def basename_no_ext(path):
    return os.path.splitext(os.path.basename(path))[0]


def build_pdf_basename_map(folder, recursive=True):
    out = {}
    for pdf_path in iter_input_pdf_files(folder, recursive=recursive):
        k = basename_no_ext(pdf_path).lower()
        out.setdefault(k, []).append(pdf_path)
    return out


def unique_dest_path(dst_dir, basename):
    dst = os.path.join(dst_dir, basename)
    if not os.path.exists(dst):
        return dst
    name, ext = os.path.splitext(basename)
    ts = time.strftime("%Y%m%d_%H%M%S")
    cand = os.path.join(dst_dir, "%s_%s%s" % (name, ts, ext))
    if not os.path.exists(cand):
        return cand
    idx = 1
    while True:
        cand2 = os.path.join(dst_dir, "%s_%s_%d%s" % (name, ts, idx, ext))
        if not os.path.exists(cand2):
            return cand2
        idx += 1


def apply_paths_to_module(mod, dest_dir, archive_dir, out_dir1, out_dir2):
    """给 get_best6 注入路径。"""
    if hasattr(mod, "set_runtime_paths"):
        try:
            mod.set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2)
            return
        except Exception:
            pass

    if hasattr(mod, "DEST_DIR"):
        mod.DEST_DIR = dest_dir
    if hasattr(mod, "IN_PDF_ARCHIVE_DIR"):
        mod.IN_PDF_ARCHIVE_DIR = archive_dir
    if hasattr(mod, "DEST_DIR1"):
        mod.DEST_DIR1 = out_dir1
    if hasattr(mod, "DEST_DIR2"):
        mod.DEST_DIR2 = out_dir2

    if hasattr(mod, "OUT_PDF_P1"):
        try:
            mod.OUT_PDF_P1 = os.path.join(out_dir1, OUT_NAME_P1)
        except Exception:
            pass
    if hasattr(mod, "OUT_PDF_P2"):
        try:
            mod.OUT_PDF_P2 = os.path.join(out_dir2, OUT_NAME_P2)
        except Exception:
            pass
    if hasattr(mod, "OUT_PDF"):
        try:
            mod.OUT_PDF = os.path.join(out_dir1, "over_test.pdf")
        except Exception:
            pass


# -------------------------
# ✅ 源PDF -> AI目录(复制文件并改后缀)
# -------------------------
def copy_pdf_to_ai_dir(src_pdf_dir, ai_dir, recursive=True, overwrite=False, log_cb=None, stop_flag=None):
    """
    把 src_pdf_dir 里的 PDF 复制到 ai_dir，并改后缀为 .ai
    注意：
    - 不是格式转换，只是复制文件内容并改扩展名
    - 保留源文件夹不变
    - 尽量保持相对目录结构
    """
    if not os.path.isdir(src_pdf_dir):
        if log_cb:
            log_cb("[FATAL] 源文件夹不存在: %s\n" % src_pdf_dir)
        return {"ok": 0, "skip": 0, "reuse": 0, "err": 1, "pairs": []}

    ensure_dir(ai_dir)

    cnt_ok = 0
    cnt_reuse = 0
    cnt_err = 0
    pairs = []

    for pdf_path in iter_input_pdf_files(src_pdf_dir, recursive=recursive):
        if stop_flag and stop_flag():
            break

        try:
            rel = os.path.relpath(pdf_path, src_pdf_dir)
            rel_no_ext = os.path.splitext(rel)[0] + ".ai"
            ai_path = os.path.join(ai_dir, rel_no_ext)

            ai_parent = os.path.dirname(ai_path)
            if ai_parent:
                ensure_dir(ai_parent)

            if os.path.exists(ai_path):
                if overwrite:
                    os.remove(ai_path)
                else:
                    if log_cb:
                        log_cb("[REUSE exists] %s\n" % ai_path)
                    cnt_reuse += 1
                    pairs.append({
                        "src_pdf": pdf_path,
                        "ai_path": ai_path
                    })
                    continue

            shutil.copy2(pdf_path, ai_path)
            if log_cb:
                log_cb("[OK] %s -> %s\n" % (pdf_path, ai_path))
            cnt_ok += 1
            pairs.append({
                "src_pdf": pdf_path,
                "ai_path": ai_path
            })

        except Exception as e:
            if log_cb:
                log_cb("[ERR] %s  %s\n" % (pdf_path, repr(e)))
            cnt_err += 1

    if log_cb:
        log_cb("[DONE copy pdf->ai] OK=%d REUSE=%d ERR=%d\n" % (cnt_ok, cnt_reuse, cnt_err))

    return {
        "ok": cnt_ok,
        "skip": cnt_reuse,
        "reuse": cnt_reuse,
        "err": cnt_err,
        "pairs": pairs
    }


# -------------------------
# ✅ 运行 JSX/VBS 批处理 AI
# -------------------------
def run_jsx_batch(ai_files, cx, vbs_path, jsx_path, out_dir, log_cb=None, stop_flag=None, proc_setter=None):
    """
    对给定 ai_files 中每个 .ai 执行：
      cscript //nologo run_ai.vbs AItest3.jsx "2;AI全路径;输出目录"

    返回：
    {
        "rc": 0/5,
        "ok_ai": [...],
        "fail_ai": [...],
        "ok_count": n,
        "fail_count": n,
        "total": n
    }
    """
    if not os.path.isfile(vbs_path):
        raise RuntimeError("找不到 vbs：%s" % vbs_path)
    if not os.path.isfile(jsx_path):
        raise RuntimeError("找不到 jsx：%s" % jsx_path)

    ai_files = [p for p in (ai_files or []) if p and os.path.isfile(p)]
    if not ai_files:
        if log_cb:
            log_cb("⚠️ 没有可处理的 .ai 文件。\n")
        return {
            "rc": 0,
            "ok_ai": [],
            "ok_pdf": [],
            "fail_ai": [],
            "ok_count": 0,
            "fail_count": 0,
            "total": 0
        }

    total = len(ai_files)
    ok_count = 0
    fail_count = 0
    ok_ai = []
    ok_pdf = []
    ok_pdf_seen = set()
    fail_ai = []

    for i, ai_path in enumerate(ai_files, start=1):
        if stop_flag and stop_flag():
            break

        ai_parent_dir = os.path.dirname(os.path.abspath(ai_path))
        target_out_dir = os.path.abspath(out_dir) if out_dir else ai_parent_dir

        probe_dirs = []
        for probe_dir in (target_out_dir, ai_parent_dir):
            if not probe_dir:
                continue
            if norm_path(probe_dir) in {norm_path(p) for p in probe_dirs}:
                continue
            probe_dirs.append(probe_dir)

        before_pdf = snapshot_pdf_mtimes_multi(probe_dirs)
        data = "%s;%s;%s" % (str(cx), ai_path, target_out_dir)
        cmd = ["cscript", "//nologo", vbs_path, jsx_path, data]

        if log_cb:
            log_cb("RUN -> %s (%d/%d)\n" % (data, i, total))
            log_cb("CMD: %s\n" % " ".join(cmd))

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        jsx_stream_encoding = "mbcs" if os.name == "nt" else "utf-8"

        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding=jsx_stream_encoding,
            errors="replace",
            bufsize=1,
            universal_newlines=True,
            env=env
        )
        if proc_setter:
            proc_setter(p)

        ret_failed = False
        ret_payload = ""

        try:
            for line in p.stdout:
                if log_cb:
                    log_cb(line)

                s = line.strip()
                if s.startswith("RET:"):
                    ret_payload = s[4:].strip()
                    up = ret_payload.upper()
                    # JSX 脚本约定：500;... 为业务失败
                    if up.startswith("500;") or up.startswith("ERR"):
                        ret_failed = True

                if stop_flag and stop_flag():
                    try:
                        if p.poll() is None:
                            p.terminate()
                    except Exception:
                        pass
                    break
        finally:
            try:
                p.stdout.close()
            except Exception:
                pass

        try:
            rc = p.wait(timeout=10)
        except Exception:
            try:
                p.kill()
            except Exception:
                pass
            rc = -1

        if proc_setter:
            proc_setter(None)

        if rc != 0:
            fail_count += 1
            fail_ai.append(ai_path)
            if log_cb:
                log_cb("❌ JSX 子进程失败 rc=%s  file=%s\n" % (rc, ai_path))
                log_cb("⚠️ 跳过失败文件，继续后续文件。\n")
            continue

        if ret_failed:
            fail_count += 1
            fail_ai.append(ai_path)
            if log_cb:
                log_cb("⚠️ JSX 返回失败，已跳过：file=%s  ret=%s\n" % (ai_path, ret_payload))
            continue

        ok_count += 1
        ok_ai.append(ai_path)

        after_pdf = snapshot_pdf_mtimes_multi(probe_dirs)
        produced_pdf = []
        for nk, info in after_pdf.items():
            p_after, mt_after, sz_after = info
            old = before_pdf.get(nk)
            if (old is None) or (mt_after > old[1] + 1e-6) or (sz_after != old[2]):
                produced_pdf.append(p_after)
        produced_pdf.sort()

        # Fallback: conventional output file name (same basename as AI)
        if not produced_pdf:
            expected_name = os.path.splitext(os.path.basename(ai_path))[0] + ".pdf"
            for probe_dir in probe_dirs:
                expected_pdf = os.path.join(probe_dir, expected_name)
                if os.path.isfile(expected_pdf):
                    produced_pdf.append(expected_pdf)

        for p_out in produced_pdf:
            nk = norm_path(p_out)
            if nk in ok_pdf_seen:
                continue
            ok_pdf_seen.add(nk)
            ok_pdf.append(p_out)

        if log_cb:
            if produced_pdf:
                log_cb("JSX导出PDF=%d file=%s\n" % (len(produced_pdf), ai_path))
            else:
                log_cb("⚠️ JSX成功但未检测到导出PDF：%s\n" % ai_path)

        # 兼容原来的进度机制
        if log_cb:
            log_cb("PROGRESS: %d / %d\n" % (i, total))

    if log_cb:
        log_cb("JSX批处理统计：成功=%d 失败=%d 总数=%d\n" % (ok_count, fail_count, total))

    rc_final = 0
    if total > 0 and ok_count == 0 and fail_count > 0:
        rc_final = 5

    return {
        "rc": rc_final,
        "ok_ai": ok_ai,
        "ok_pdf": ok_pdf,
        "fail_ai": fail_ai,
        "ok_count": ok_count,
        "fail_count": fail_count,
        "total": total
    }


# -------------------------
# CLI 子进程模式：仅运行 get_best6
# -------------------------
def _cli_run_algo(input_dir, work_dir, out1, out2, disable_single_single_line_special_mode=False):
    # PDF 直接在 input_dir 消费，不再复制/迁移到 work_dir
    print("=== RUN get_best6.py (拼图) ===")
    print("ALGO_INPUT_DIR:", input_dir)
    print("ALGO_ARCHIVE_DIR:", input_dir)
    print("SINGLE_SINGLE_LINE_SPECIAL_MODE:", "OFF" if disable_single_single_line_special_mode else "ON")
    res = get_best6.run(
        cfg={
            "DEST_DIR": input_dir,
            "TEST_DIR": input_dir,
            "DEST_DIR1": out1,
            "DEST_DIR2": out2,
            "SINGLE_SINGLE_LINE_SPECIAL_MODE": (not bool(disable_single_single_line_special_mode)),
        },
        input_pdfs=None,
        progress_cb=None,
        log_cb=None,
    )
    print(
        "DONE: single_p1=%d single_p2=%d mix_p1=%d mix_p2=%d"
        % (
            len(res.get("single_p1_files") or []),
            len(res.get("single_p2_files") or []),
            len(res.get("mix_p1_files") or []),
            len(res.get("mix_p2_files") or []),
        )
    )
    return 0


def _is_cli_mode():
    return ("--run-algo" in sys.argv)


def _cli_entry():
    # 简单解析参数(避免引入 argparse)
    if "--run-algo" in sys.argv:
        def _get(flag, default=""):
            if flag in sys.argv:
                j = sys.argv.index(flag)
                return sys.argv[j + 1] if j + 1 < len(sys.argv) else default
            return default

        work_dir = _get("--work", "")
        input_dir = _get("--input", "")
        out1 = _get("--out1", "")
        out2 = _get("--out2", "")
        disable_single_single_line_special_mode = ("--disable-single-single-line-special-mode" in sys.argv)

        if not input_dir:
            input_dir = work_dir

        if not input_dir or not work_dir or not out1 or not out2:
            print("ERR: missing args for --run-algo")
            return 2

        try:
            return _cli_run_algo(
                input_dir,
                work_dir,
                out1,
                out2,
                disable_single_single_line_special_mode=disable_single_single_line_special_mode,
            )
        except Exception:
            print("❌ ALGO EXCEPTION:\n" + traceback.format_exc())
            return 1

    return 0


# -------------------------
# Worker Thread
# -------------------------
class RunnerThread(QThread):
    sig_log = Signal(str)
    sig_status = Signal(str, str)   # status, phase
    sig_progress = Signal(float)    # 0..100
    sig_done = Signal(int)          # rc

    def __init__(self, cfg, parent=None):
        super(RunnerThread, self).__init__(parent)
        self.cfg = cfg
        self._stop = False
        self.current_proc = None

    def request_stop(self):
        self._stop = True
        self.sig_log.emit("\n⛔ 请求停止...\n")
        try:
            if self.current_proc and (self.current_proc.poll() is None):
                self.current_proc.terminate()
                self.sig_log.emit("⛔ 已发送 terminate 给子进程\n")
        except Exception:
            pass

    def _set_current_proc(self, p):
        self.current_proc = p

    def _spawn_and_stream(self, cmd):
        self.sig_log.emit("CMD: %s\n" % " ".join(cmd))

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            universal_newlines=True,
            env=env
        )
        self.current_proc = p

        try:
            for line in p.stdout:
                self.sig_log.emit(line)

                # 解析进度：PROGRESS: a / b
                s = line.strip()
                m = re.search(r"PROGRESS:\s*(\d+)\s*/\s*(\d+)", s)
                if m:
                    cur = float(m.group(1))
                    tot = float(m.group(2)) if float(m.group(2)) > 0 else 1.0
                    pct = max(0.0, min(100.0, cur * 100.0 / tot))
                    self.sig_progress.emit(pct)

                if self._stop:
                    try:
                        if p.poll() is None:
                            p.terminate()
                    except Exception:
                        pass
                    break
        finally:
            try:
                p.stdout.close()
            except Exception:
                pass

        try:
            rc = p.wait(timeout=5)
        except Exception:
            try:
                p.kill()
            except Exception:
                pass
            rc = -1

        self.current_proc = None
        return rc

    def _backup_matching_source_files(self, exported_files, source_dir, backup_dir):
        ensure_dir(backup_dir)

        source_map = build_pdf_basename_map(source_dir, recursive=True)
        backed_up = []
        seen_source = set()
        no_match = 0

        for item in exported_files:
            if self._stop:
                break

            if isinstance(item, dict):
                ref_path = (item or {}).get("src") or (item or {}).get("dst") or ""
            else:
                ref_path = str(item or "")
            if not ref_path:
                continue

            key = basename_no_ext(ref_path).lower()
            matches = source_map.get(key, [])
            if not matches:
                no_match += 1
                self.sig_log.emit("⚠️ 未找到同名源文件，跳过备份：%s\n" % os.path.basename(ref_path))
                continue

            for src_path in matches:
                nk = norm_path(src_path)
                if nk in seen_source:
                    continue
                seen_source.add(nk)

                dst_path = unique_dest_path(backup_dir, os.path.basename(src_path))
                shutil.copy2(src_path, dst_path)
                backed_up.append(dst_path)
                self.sig_log.emit("[BACKUP] %s -> %s\n" % (src_path, dst_path))

        self.sig_log.emit("同名源文件备份完成：%d 个\n" % len(backed_up))
        if no_match:
            self.sig_log.emit("未匹配到同名源文件：%d 个\n" % no_match)

        return backed_up

    def _run_separation_pipeline(self, source_dir, ai_dir, source_backup_batch_dir, cx, jsx_path, vbs_path):
        # 1) 源文件 -> AI目录(复制并改后缀)
        self.sig_status.emit("Running", "PDF->AI")
        self.sig_log.emit("=== STEP1: COPY SOURCE PDF -> AI DIR (.ai extension) ===\n")
        prep = copy_pdf_to_ai_dir(
            src_pdf_dir=source_dir,
            ai_dir=ai_dir,
            recursive=True,
            overwrite=False,
            log_cb=lambda s: self.sig_log.emit(s),
            stop_flag=lambda: self._stop
        )
        if self._stop:
            self.sig_log.emit("⛔ 已停止：不再继续。\n")
            return {"abort": True, "rc": 0, "ok_pdf": []}

        pairs = prep.get("pairs", [])
        if not pairs:
            self.sig_log.emit("⚠️ 源文件夹中没有可处理的 PDF。\n")
            return {"abort": False, "rc": 0, "ok_pdf": []}

        ai_files_to_run = [x["ai_path"] for x in pairs if x.get("ai_path")]

        # 2) 执行 JSX 批处理
        self.sig_status.emit("Running", "AI->PDF (JSX)")
        self.sig_log.emit("\n=== STEP2: RUN JSX (AI -> PDF export) ===\n")
        jsx_res = run_jsx_batch(
            ai_files=ai_files_to_run,
            cx=cx,
            vbs_path=vbs_path,
            jsx_path=jsx_path,
            out_dir=ai_dir,
            log_cb=lambda s: self.sig_log.emit(s),
            stop_flag=lambda: self._stop,
            proc_setter=self._set_current_proc
        )
        if self._stop:
            self.sig_log.emit("⛔ 已停止：JSX阶段中断。\n")
            return {"abort": True, "rc": 0, "ok_pdf": []}

        rc_jsx = jsx_res.get("rc", 0)
        if rc_jsx != 0:
            self.sig_log.emit("❌ JSX 执行失败 rc=%s\n" % rc_jsx)
            return {"abort": True, "rc": rc_jsx, "ok_pdf": []}

        ok_ai = [p for p in (jsx_res.get("ok_ai", []) or []) if p and os.path.isfile(p)]
        ok_pdf = [p for p in (jsx_res.get("ok_pdf", []) or []) if p and os.path.isfile(p)]
        if not ok_pdf:
            self.sig_log.emit("\n⚠️ 没有检测到 JSX 成功导出的 PDF。\n")
            return {"abort": False, "rc": 0, "ok_pdf": []}

        self.sig_log.emit("JSX成功AI数量：%d\n" % len(ok_ai))
        self.sig_log.emit("JSX导出PDF数量：%d\n" % len(ok_pdf))
        self.sig_log.emit("JSX导出PDF目录：%s\n" % ai_dir)

        ai_dir_norm = norm_path(ai_dir)
        not_in_ai_root = [
            p for p in ok_pdf
            if norm_path(os.path.dirname(os.path.abspath(p))) != ai_dir_norm
        ]
        if not_in_ai_root:
            self.sig_log.emit("⚠️ 检测到部分导出 PDF 没有落在转化文件夹根目录，已自动跳过：\n")
            for p in not_in_ai_root[:20]:
                self.sig_log.emit("   %s\n" % p)
            if len(not_in_ai_root) > 20:
                self.sig_log.emit("   ... 还有 %d 个\n" % (len(not_in_ai_root) - 20))
            ok_pdf = [p for p in ok_pdf if p not in not_in_ai_root]
            if not ok_pdf:
                self.sig_log.emit("⚠️ 跳过后没有可继续处理的 PDF。\n")
                return {"abort": False, "rc": 0, "ok_pdf": []}

        # 3) 备份与本次导出 PDF 同名的源文件
        self.sig_status.emit("Running", "Backup")
        self.sig_log.emit("\n=== STEP3: BACKUP MATCHING SOURCE FILES ===\n")
        backed_up = self._backup_matching_source_files(
            ok_pdf,
            source_dir,
            source_backup_batch_dir
        )
        if not backed_up:
            self.sig_log.emit("⚠️ 本次导出未匹配到需要备份的源文件。\n")
        else:
            self.sig_log.emit("源文件备份目录：%s\n" % source_backup_batch_dir)
        self.sig_log.emit("\n")
        if self._stop:
            self.sig_log.emit("⛔ 已停止：不再继续。\n")
            return {"abort": True, "rc": 0, "ok_pdf": ok_pdf}

        return {"abort": False, "rc": 0, "ok_pdf": ok_pdf}

    def _run_algorithm_pipeline(self, ai_dir, out1, out2):
        self.sig_status.emit("Running", "Algorithm")
        if is_frozen_app():
            self.sig_log.emit("=== STEP4: RUN ALGO (INLINE/FROZEN) ===\n")
            if self._stop:
                self.sig_log.emit("⛔ 已停止：算法阶段未开始。\n")
                return {"abort": True, "rc": 0}

            def _algo_log_cb(msg):
                try:
                    self.sig_log.emit(str(msg).rstrip("\n") + "\n")
                except Exception:
                    pass

            def _algo_progress_cb(cur, tot, msg):
                try:
                    cur_f = float(cur)
                    tot_f = float(tot) if float(tot) > 0 else 1.0
                except Exception:
                    cur_f, tot_f = 0.0, 1.0
                pct = max(0.0, min(100.0, cur_f * 100.0 / tot_f))
                self.sig_progress.emit(pct)
                phase = ("Algorithm: %s" % msg) if msg else "Algorithm"
                self.sig_status.emit("Running", phase)
                if msg:
                    self.sig_log.emit("PROGRESS: %s / %s  %s\n" % (int(cur_f), int(tot_f), msg))

            try:
                res = get_best6.run(
                    cfg={
                        "DEST_DIR": ai_dir,
                        "TEST_DIR": ai_dir,
                        "DEST_DIR1": out1,
                        "DEST_DIR2": out2,
                        "SINGLE_SINGLE_LINE_SPECIAL_MODE": (not bool(self.cfg.get("disable_single_single_line_special_mode", False))),
                    },
                    input_pdfs=None,
                    progress_cb=_algo_progress_cb,
                    log_cb=_algo_log_cb
                )
                self.sig_progress.emit(100.0)
                self.sig_log.emit(
                    "DONE: single_p1=%d single_p2=%d mix_p1=%d mix_p2=%d\n" % (
                        len(res.get("single_p1_files") or []),
                        len(res.get("single_p2_files") or []),
                        len(res.get("mix_p1_files") or []),
                        len(res.get("mix_p2_files") or []),
                    )
                )
                return {"abort": False, "rc": 0}
            except Exception:
                self.sig_log.emit("\n❌ 算法直跑失败:\n" + traceback.format_exc() + "\n")
                return {"abort": True, "rc": 1}

        self.sig_log.emit("=== STEP4: RUN ALGO (SUBPROCESS) ===\n")
        cmd_algo = [
            sys.executable, "-u", os.path.abspath(__file__),
            "--run-algo",
            "--input", ai_dir,
            "--work", ai_dir,
            "--out1", out1,
            "--out2", out2
        ]
        if bool(self.cfg.get("disable_single_single_line_special_mode", False)):
            cmd_algo.append("--disable-single-single-line-special-mode")
        rc = self._spawn_and_stream(cmd_algo)
        if self._stop:
            self.sig_log.emit("⛔ 已停止：算法阶段中断。\n")
            return {"abort": True, "rc": 0}
        if rc != 0:
            self.sig_log.emit("\n❌ 算法子进程失败 (rc=%s)\n" % rc)
            return {"abort": True, "rc": rc}
        return {"abort": False, "rc": 0}

    def run(self):
        try:
            cfg = self.cfg
            source_dir = cfg["source_dir"]     # ✅ 第1个目录：源文件夹
            ai_dir = cfg["ai_dir"]             # ✅ 第2个目录：AI目录
            source_backup_dir = cfg["source_backup_dir"]  # ✅ 第3个目录：源文件备份
            out1 = cfg["out1"]
            out2 = cfg["out2"]
            cx = cfg.get("cx", 2)
            run_mode = cfg.get("run_mode", RUN_MODE_FULL)
            run_mode_text = RUN_MODE_TEXT.get(run_mode, run_mode)

            script_dir = os.path.dirname(os.path.abspath(__file__))
            jsx_path = os.path.join(script_dir, "run", "AItest3.jsx")
            if not os.path.isfile(jsx_path):
                jsx_path = os.path.join(script_dir, "AItest3.jsx")
            vbs_path = os.path.join(script_dir, "run_ai.vbs")

            self.sig_status.emit("Running", "Preparing")
            self.sig_progress.emit(0.0)

            if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE):
                ensure_dir(ai_dir)
                ensure_dir(source_backup_dir)
            if run_mode in (RUN_MODE_FULL, RUN_MODE_ALGO):
                ensure_dir(out1)
                ensure_dir(out2)

            if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE):
                batch_name = make_unique_batch_name([source_backup_dir])
                source_backup_batch_dir = os.path.join(source_backup_dir, batch_name)
                ensure_dir(source_backup_batch_dir)
            else:
                batch_name = ""
                source_backup_batch_dir = ""

            self.sig_log.emit("=== CONFIG ===\n")
            self.sig_log.emit("运行模式 RUN_MODE                    : %s\n" % run_mode_text)
            self.sig_log.emit("源文件夹 SOURCE_DIR                   : %s\n" % source_dir)
            self.sig_log.emit("AI目录 AI_DIR                        : %s\n" % ai_dir)
            self.sig_log.emit("源文件备份 SOURCE_BACKUP_DIR          : %s\n" % source_backup_dir)
            self.sig_log.emit("源文件备份(本次) SOURCE_BACKUP_BATCH   : %s\n" % (source_backup_batch_dir or "（当前模式未使用）"))
            self.sig_log.emit("本次批次 RUN_BATCH_NAME              : %s\n" % (batch_name or "（当前模式未使用）"))
            self.sig_log.emit("输出拼图 DEST_DIR1                   : %s\n" % out1)
            self.sig_log.emit("输出刀线 DEST_DIR2                   : %s\n" % out2)
            self.sig_log.emit("JSX                                  : %s\n" % jsx_path)
            self.sig_log.emit("VBS                                  : %s\n" % vbs_path)
            self.sig_log.emit("CX                                   : %s\n" % str(cx))
            self.sig_log.emit(
                "单枚单线特殊模式 SINGLE_SINGLE_LINE_SPECIAL_MODE : %s\n"
                % ("关闭" if bool(cfg.get("disable_single_single_line_special_mode", False)) else "开启")
            )
            self.sig_log.emit("运行算法 ALGO                         : get_best6 (拼图)\n")
            self.sig_log.emit("================\n\n")

            if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE) and (not os.path.isdir(source_dir)):
                self.sig_log.emit("❌ 源文件夹不存在：%s\n" % source_dir)
                self.sig_done.emit(2)
                return

            if run_mode == RUN_MODE_ALGO and (not os.path.isdir(ai_dir)):
                self.sig_log.emit("❌ 转化文件夹不存在：%s\n" % ai_dir)
                self.sig_done.emit(2)
                return

            if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE):
                sep_res = self._run_separation_pipeline(
                    source_dir=source_dir,
                    ai_dir=ai_dir,
                    source_backup_batch_dir=source_backup_batch_dir,
                    cx=cx,
                    jsx_path=jsx_path,
                    vbs_path=vbs_path
                )
                if sep_res.get("abort"):
                    self.sig_done.emit(int(sep_res.get("rc", 0)))
                    return
                if run_mode == RUN_MODE_SEPARATE:
                    self.sig_log.emit("✅ 当前模式已完成：仅分离\n")
                    self.sig_done.emit(0)
                    return
                if not (sep_res.get("ok_pdf") or []):
                    self.sig_log.emit("⚠️ 分离阶段没有得到可用于拼接的 PDF，本次一气呵成到此结束。\n")
                    self.sig_done.emit(0)
                    return

            if run_mode in (RUN_MODE_FULL, RUN_MODE_ALGO):
                algo_res = self._run_algorithm_pipeline(ai_dir=ai_dir, out1=out1, out2=out2)
                if algo_res.get("abort"):
                    self.sig_done.emit(int(algo_res.get("rc", 0)))
                    return

            self.sig_done.emit(0)

        except Exception:
            self.sig_log.emit("\n❌ RUNNER EXCEPTION:\n" + traceback.format_exc() + "\n")
            self.sig_done.emit(1)


# -------------------------
# Help Dialog
# -------------------------
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super(HelpDialog, self).__init__(parent)
        self.setWindowTitle("说明：源PDF -> AI预处理 + 拼图运行器")
        self.resize(920, 580)

        layout = QVBoxLayout(self)
        self.txt = QTextEdit(self)
        self.txt.setReadOnly(True)
        self.txt.setFont(QFont("Consolas", 11))
        layout.addWidget(self.txt)

        btn = QPushButton("返回", self)
        btn.clicked.connect(self.accept)
        layout.addWidget(btn, alignment=Qt.AlignRight)

        help_text = (
            "【功能模式】\n"
            "1)仅分离：源文件 -> AI目录(.ai) -> JSX导出PDF -> 同名源文件备份\n"
            "2)仅拼接：直接读取“转化文件夹”中的 PDF，运行 get_best6 输出印刷/刀版文件\n"
            "3)一气呵成：按完整流程连续执行分离 + 拼接\n\n"
            "【目录说明(保持不变)】\n"
            "1)第1个目录选择：源文件夹(原始PDF)\n"
            "2)第2个目录选择：AI目录\n"
            "3)第3个目录选择：源文件  (备份)(程序会在这里自动新建同名时间子目录)\n"
            "4)程序会把第1个目录中的 PDF 复制到第2个目录，并改后缀为 .ai(仅改扩展名，不是格式转换)\n"
            "5)对这些 .ai 执行：cscript //nologo run_ai.vbs AItest3.jsx \"2;AI全路径;输出目录\"\n"
            "6)JSX 导出的 PDF 会直接保存到第2个目录根层；同名文件自动覆盖，不再弹确认框\n"
            "7)把与本次导出 PDF 同名的源文件复制到“源文件  (备份)”下同名时间文件夹\n"
            "8)子进程直接读取 AI 目录里的 PDF，运行 get_best6 进行拼图输出\n\n"
            "【注意】\n"
            "- run.py、AItest3.jsx、run_ai.vbs 必须在同一目录\n"
            "- get_best6 现在直接读取 AI 目录中的 PDF\n"
            "- 源文件  (备份) 只备份与本次导出 PDF 同名的源 PDF\n"
            "- 单个 AI 文件如果分离失败，会自动跳过并继续后面的文件\n"
        )
        self.txt.setPlainText(help_text)


# -------------------------
# Main Window
# -------------------------
class RollWidget(QWidget):
    def __init__(self):
        super(RollWidget, self).__init__()

        self.setWindowTitle("印客链 - 源PDF -> AI预处理 + 拼图运行器")

        ico_path = r"C:\Users\wzqy\PycharmProjects\inklink\icon.ico"
        if os.path.isfile(ico_path):
            try:
                self.setWindowIcon(QIcon(ico_path))
            except Exception:
                pass

        self.resize(1120, 800)
        self.worker = None
        self._loading_settings = False
        self.settings = QSettings("InkLink", "PdfAiRunner")

        self._build_ui()
        self._apply_qss()

        # 默认值
        self.ed_src_pdf.setText(r"D:\我的数据\文档\get")
        self.ed_ai.setText(r"D:\我的数据\文档\已解密文件")
        self.ed_src_backup.setText(r"D:\我的数据\文档\源文件  (备份)")
        self.ed_out1.setText(r"D:\我的数据\文档\2")
        self.ed_out2.setText(r"D:\我的数据\文档\3")

        self._load_settings()
        self._bind_settings_signals()

        self._set_status("Idle", "Ready")
        self._set_progress_indeterminate(False)
        self.clear_log()

    def _bind_settings_signals(self):
        self.ed_src_pdf.textChanged.connect(self._save_settings)
        self.ed_ai.textChanged.connect(self._save_settings)
        self.ed_src_backup.textChanged.connect(self._save_settings)
        self.ed_out1.textChanged.connect(self._save_settings)
        self.ed_out2.textChanged.connect(self._save_settings)
        self.chk_disable_single_single_line_special_mode.toggled.connect(self._save_settings)

    def _load_settings(self):
        self._loading_settings = True
        try:
            self.ed_src_pdf.setText(self.settings.value("paths/source_dir", self.settings.value("paths/src_pdf_dir", self.ed_src_pdf.text(), type=str), type=str))
            self.ed_ai.setText(self.settings.value("paths/ai_dir", self.ed_ai.text(), type=str))
            self.ed_src_backup.setText(self.settings.value("paths/source_backup_dir", self.ed_src_backup.text(), type=str))
            self.ed_out1.setText(self.settings.value("paths/out1", self.ed_out1.text(), type=str))
            self.ed_out2.setText(self.settings.value("paths/out2", self.ed_out2.text(), type=str))
            self.chk_disable_single_single_line_special_mode.setChecked(
                bool(self.settings.value("options/disable_single_single_line_special_mode", False, type=bool))
            )
        finally:
            self._loading_settings = False

    def _save_settings(self, *args):
        if self._loading_settings:
            return

        source_dir = self.ed_src_pdf.text().strip()

        self.settings.setValue("paths/source_dir", source_dir)
        self.settings.setValue("paths/src_pdf_dir", source_dir)
        self.settings.setValue("paths/ai_dir", self.ed_ai.text().strip())
        self.settings.setValue("paths/source_backup_dir", self.ed_src_backup.text().strip())
        self.settings.setValue("paths/out1", self.ed_out1.text().strip())
        self.settings.setValue("paths/out2", self.ed_out2.text().strip())
        self.settings.setValue(
            "options/disable_single_single_line_special_mode",
            bool(self.chk_disable_single_single_line_special_mode.isChecked()),
        )
        self.settings.sync()

    def _apply_qss(self):
        qss = """
        QWidget { background: #0f172a; color: #e5e7eb; font-size: 12px; }
        QGroupBox {
            border: 1px solid #23304f;
            margin-top: 10px;
            border-radius: 8px;
            padding: 10px;
            background: #0b1224;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px 0 6px;
            color: #e5e7eb;
            font-weight: 600;
        }
        QLineEdit {
            background: #101b33;
            border: 1px solid #23304f;
            border-radius: 6px;
            padding: 8px;
        }
        QPushButton {
            background: #101b33;
            border: 1px solid #23304f;
            border-radius: 8px;
            padding: 10px 14px;
        }
        QPushButton:hover { background: #1f2a44; }
        QPushButton#btnStart {
            background: #3b82f6;
            border: none;
            font-weight: 700;
        }
        QPushButton#btnStart:hover { background: #2563eb; }
        QPushButton#btnSeparate {
            background: #0f766e;
            border: none;
            font-weight: 700;
        }
        QPushButton#btnSeparate:hover { background: #0d9488; }
        QPushButton#btnAlgo {
            background: #f59e0b;
            color: #111827;
            border: none;
            font-weight: 700;
        }
        QPushButton#btnAlgo:hover { background: #d97706; }
        QPushButton#btnStop {
            background: #ef4444;
            border: none;
            font-weight: 700;
        }
        QPushButton#btnStop:hover { background: #dc2626; }
        QTextEdit {
            background: #0b1224;
            border: 1px solid #23304f;
            border-radius: 8px;
            padding: 10px;
        }
        QProgressBar {
            background: #101b33;
            border: 1px solid #23304f;
            border-radius: 6px;
            text-align: center;
        }
        QProgressBar::chunk { background: #3b82f6; border-radius: 6px; }
        """
        self.setStyleSheet(qss)

    def _build_ui(self):
        root = QVBoxLayout()
        self.setLayout(root)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        # Header
        header = QFrame(self)
        hl = QHBoxLayout(header)
        hl.setContentsMargins(12, 12, 12, 12)
        hl.setSpacing(12)

        title_box = QVBoxLayout()
        lb_title = QLabel("印客链 - 源PDF -> AI预处理 + 拼图运行器")
        f = QFont("Microsoft YaHei UI", 18)
        f.setBold(True)
        lb_title.setFont(f)
        lb_sub = QLabel("支持：仅分离 / 仅拼接 / 一气呵成，目录配置保持不变")
        lb_sub.setStyleSheet("color:#94a3b8;")
        title_box.addWidget(lb_title)
        title_box.addWidget(lb_sub)
        hl.addLayout(title_box, 1)

        st_box = QVBoxLayout()
        self.lb_status = QLabel("Idle")
        fs = QFont("Microsoft YaHei UI", 12)
        fs.setBold(True)
        self.lb_status.setFont(fs)
        self.lb_phase = QLabel("Ready")
        self.lb_phase.setStyleSheet("color:#94a3b8;")
        st_box.addWidget(self.lb_status, alignment=Qt.AlignRight)
        st_box.addWidget(self.lb_phase, alignment=Qt.AlignRight)
        hl.addLayout(st_box, 0)

        self.pb = QProgressBar()
        self.pb.setFixedWidth(280)
        self.pb.setRange(0, 100)
        self.pb.setValue(0)
        hl.addWidget(self.pb, 0, alignment=Qt.AlignVCenter)

        root.addWidget(header)

        splitter = QSplitter(Qt.Horizontal, self)
        root.addWidget(splitter, 1)

        # Left panel
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(12)

        gp_paths = QGroupBox("路径配置")
        grid = QGridLayout(gp_paths)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        def add_path_row(r, label, pick=True):
            lb = QLabel(label)
            ed = QLineEdit()
            btn = QPushButton("📁 选择") if pick else QPushButton("-")
            if pick:
                btn.clicked.connect(lambda: self._pick_dir(ed))
            else:
                btn.setEnabled(False)
            grid.addWidget(lb, r, 0)
            grid.addWidget(ed, r, 1)
            grid.addWidget(btn, r, 2)
            return ed

        # ✅ 第1个目录：源文件夹
        self.ed_src_pdf = add_path_row(0, "源文件夹  (原始PDF)")

        # ✅ 第2个目录：AI目录
        self.ed_ai = add_path_row(1, "转化文件夹 (PDF文件转化为AI文件)")

        # ✅ 第3个目录：源文件备份
        self.ed_src_backup = add_path_row(2, "备份文件夹  (源文件, 自动创建时间子目录)")

        # ✅ 第4~5个保持不变
        self.ed_out1 = add_path_row(3, "印刷文件夹 (输出PDF 印刷)")
        self.ed_out2 = add_path_row(4, "刀版文件夹 (输出刀版文件)")

        grid.setColumnStretch(1, 1)
        left_layout.addWidget(gp_paths)

        gp_mode = QGroupBox("运行功能")
        mode_layout = QVBoxLayout(gp_mode)
        lb_mode = QLabel("目录不变，可按需要单独运行“分离”或“拼接”，也可继续直接一气呵成。")
        lb_mode.setStyleSheet("color:#94a3b8;")
        mode_layout.addWidget(lb_mode)

        mode_row = QHBoxLayout()
        self.btn_run_separate = QPushButton("① 仅分离")
        self.btn_run_separate.setObjectName("btnSeparate")
        self.btn_run_separate.clicked.connect(self.start_separate)
        mode_row.addWidget(self.btn_run_separate)

        self.btn_run_algo = QPushButton("② 仅拼接")
        self.btn_run_algo.setObjectName("btnAlgo")
        self.btn_run_algo.clicked.connect(self.start_algo)
        mode_row.addWidget(self.btn_run_algo)

        self.btn_start = QPushButton("▶ 一气呵成")
        self.btn_start.setObjectName("btnStart")
        self.btn_start.clicked.connect(self.start)
        mode_row.addWidget(self.btn_start)
        mode_layout.addLayout(mode_row)

        self.chk_disable_single_single_line_special_mode = QCheckBox("点击后关闭：单枚单线引号内规则")
        self.chk_disable_single_single_line_special_mode.setChecked(False)
        self.chk_disable_single_single_line_special_mode.setStyleSheet("color:#cbd5e1;")
        mode_layout.addWidget(self.chk_disable_single_single_line_special_mode)
        left_layout.addWidget(gp_mode)

        gp_btn = QGroupBox("辅助操作")
        hb = QHBoxLayout(gp_btn)

        self.btn_stop = QPushButton("⛔ 停止运行")
        self.btn_stop.setObjectName("btnStop")
        self.btn_stop.clicked.connect(self.stop)
        self.btn_stop.setEnabled(False)

        self.btn_help = QPushButton("ℹ 说明")
        self.btn_help.clicked.connect(self.show_help)

        self.btn_clear = QPushButton("🧹 清空日志")
        self.btn_clear.clicked.connect(self.clear_log)

        hb.addWidget(self.btn_start)
        hb.addWidget(self.btn_stop)
        hb.addWidget(self.btn_help)
        hb.addWidget(self.btn_clear)
        left_layout.addWidget(gp_btn)

        tip = QLabel("提示：仅分离会生成并备份 PDF；仅拼接会直接读取“转化文件”中的 PDF；一气呵成保留原完整流程。")
        tip.setStyleSheet("color:#94a3b8;")
        left_layout.addWidget(tip)
        left_layout.addStretch(1)

        splitter.addWidget(left)

        # Right panel (log)
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        gp_log = QGroupBox("运行日志(子进程 stdout 实时回传)")
        vlog = QVBoxLayout(gp_log)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Consolas", 11))
        vlog.addWidget(self.log)

        right_layout.addWidget(gp_log)
        splitter.addWidget(right)

        splitter.setSizes([460, 660])

    def _pick_dir(self, line_edit: QLineEdit):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹", line_edit.text().strip() or os.getcwd())
        if d:
            line_edit.setText(d)
            self._save_settings()

    def _set_status(self, status, phase):
        self.lb_status.setText(status)
        self.lb_phase.setText(phase)

    def _set_progress_indeterminate(self, on: bool):
        if on:
            self.pb.setRange(0, 0)
        else:
            self.pb.setRange(0, 100)

    @staticmethod
    def _html_escape(s):
        return (s.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;"))

    def clear_log(self):
        self.log.clear()
        self.log.append('<span style="color:#94a3b8">Ready.</span>')

    def show_help(self):
        dlg = HelpDialog(self)
        dlg.exec()

    def _append_log(self, s: str):
        line = s.rstrip("\n")
        if not line:
            self.log.append("")
            return

        color = "#e5e7eb"
        if line.startswith("CMD:"):
            color = "#60a5fa"
        elif ("❌" in line) or line.startswith("ERR:") or line.startswith("RET: 500") or ("500;ERR;" in line) or ("EXCEPTION" in line) or line.startswith("[FATAL]") or line.startswith("[ERR]"):
            color = "#fb7185"
        elif ("⚠️" in line) or line.startswith("WARN:") or line.startswith("[SKIP") or line.startswith("[REUSE"):
            color = "#fbbf24"
        elif ("✅" in line) or ("OK" in line) or line.startswith("[OK]"):
            color = "#34d399"
        elif "⛔" in line:
            color = "#fbbf24"

        if "STEP1:" in line:
            self._set_status("Running", "PDF->AI")
            self._set_progress_indeterminate(True)
        elif "STEP2:" in line:
            self._set_status("Running", "AI->PDF (JSX)")
            self._set_progress_indeterminate(True)
        elif "STEP3:" in line:
            self._set_status("Running", "Backup")
            self._set_progress_indeterminate(True)
        elif "STEP4:" in line:
            self._set_status("Running", "Algorithm")
            self._set_progress_indeterminate(True)

        self.log.append('<span style="color:%s">%s</span>' % (color, self._html_escape(line)))

    def _set_run_buttons_enabled(self, enabled):
        self.btn_start.setEnabled(enabled)
        self.btn_run_separate.setEnabled(enabled)
        self.btn_run_algo.setEnabled(enabled)

    def start(self):
        self.start_with_mode(RUN_MODE_FULL)

    def start_separate(self):
        self.start_with_mode(RUN_MODE_SEPARATE)

    def start_algo(self):
        self.start_with_mode(RUN_MODE_ALGO)

    def start_with_mode(self, run_mode):
        if self.worker is not None:
            return

        source_dir = self.ed_src_pdf.text().strip()
        ai_dir = self.ed_ai.text().strip()
        source_backup_dir = self.ed_src_backup.text().strip()
        out1 = self.ed_out1.text().strip()
        out2 = self.ed_out2.text().strip()

        if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE) and (not os.path.isdir(source_dir)):
            QMessageBox.critical(self, "错误", "源文件夹不存在：\n" + source_dir)
            return
        if not ai_dir:
            QMessageBox.critical(self, "错误", "AI目录不能为空")
            return
        if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE) and (not source_backup_dir):
            QMessageBox.critical(self, "错误", "源文件备份目录不能为空")
            return
        if run_mode in (RUN_MODE_FULL, RUN_MODE_ALGO):
            if not out1:
                QMessageBox.critical(self, "错误", "印刷文件夹不能为空")
                return
            if not out2:
                QMessageBox.critical(self, "错误", "刀版文件夹不能为空")
                return
        if run_mode == RUN_MODE_ALGO and (not os.path.isdir(ai_dir)):
            QMessageBox.critical(self, "错误", "转化文件夹不存在：\n" + ai_dir)
            return

        if run_mode in (RUN_MODE_FULL, RUN_MODE_SEPARATE):
            ensure_dir(ai_dir)
            ensure_dir(source_backup_dir)
        if run_mode in (RUN_MODE_FULL, RUN_MODE_ALGO):
            ensure_dir(out1)
            ensure_dir(out2)

        self._set_run_buttons_enabled(False)
        self.btn_stop.setEnabled(True)
        self._set_status("Running", "Preparing")
        self._set_progress_indeterminate(True)
        self.pb.setValue(0)

        self.clear_log()
        self._append_log("模式：%s\n" % RUN_MODE_TEXT.get(run_mode, run_mode))

        cfg = {
            "source_dir": source_dir,
            "ai_dir": ai_dir,
            "source_backup_dir": source_backup_dir,
            "out1": out1,
            "out2": out2,
            "cx": 2,
            "run_mode": run_mode,
            "disable_single_single_line_special_mode": bool(self.chk_disable_single_single_line_special_mode.isChecked()),
        }

        self._save_settings()

        self.worker = RunnerThread(cfg)
        self.worker.sig_log.connect(self._append_log)
        self.worker.sig_status.connect(self._set_status)
        self.worker.sig_progress.connect(self._on_progress)
        self.worker.sig_done.connect(self._on_done)
        self.worker.start()

    @Slot(float)
    def _on_progress(self, pct):
        self._set_progress_indeterminate(False)
        self.pb.setValue(int(max(0.0, min(100.0, pct))))

    @Slot(int)
    def _on_done(self, rc):
        self._set_progress_indeterminate(False)
        self.pb.setValue(0)
        if rc == 0:
            self._set_status("Done", "Finished")
            self._append_log("\n✅ DONE\n")
        else:
            self._set_status("Failed", "Stopped/Error")
            self._append_log("\n❌ FAILED (rc=%s)\n" % rc)

        self._set_run_buttons_enabled(True)
        self.btn_stop.setEnabled(False)

        try:
            self.worker.quit()
            self.worker.wait(3000)
        except Exception:
            pass
        self.worker = None

    def stop(self):
        if self.worker is None:
            return
        self._append_log("\n⛔ 请求停止...\n")
        self._set_status("Stopping", "Terminating")
        try:
            self.worker.request_stop()
        except Exception:
            pass

    def closeEvent(self, event):
        self._save_settings()
        if self.worker is not None:
            ret = QMessageBox.question(
                self, "退出", "正在运行中，确定要退出吗？(会终止子进程)",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if ret != QMessageBox.Yes:
                event.ignore()
                return
            try:
                self.worker.request_stop()
            except Exception:
                pass
        event.accept()


# -------------------------
# main
# -------------------------
if __name__ == "__main__":
    # ✅ 子进程模式
    if _is_cli_mode():
        code = _cli_entry()
        raise SystemExit(int(code))

    # ✅ GUI 模式
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei UI", 11))

    w = RollWidget()
    w.show()

    sys.exit(app.exec())
