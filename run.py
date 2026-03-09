# -*- coding: utf-8 -*-
"""
run.py（PySide6 工程GUI版：AI预处理 + 子进程运行拼图算法 get_best5）

新增（按你的要求，且运行顺序在拼图之前）：
0) 选择“输入AI文件夹”（也是 JSX 输出 PDF 的目录，也是拼图输入 PDF 的目录）
1) 自动把该文件夹里所有 .pdf/.PDF 批量改后缀为 .ai（无需手动）
2) 自动对该文件夹里所有 .ai 执行：cscript //nologo run_ai.vbs AItest_ai.jsx "2;AI全路径;输出目录"
3) 然后再把该目录中的 PDF（排除 over_test*.pdf）复制/移动到 work 目录
4) 子进程运行 get_best5 进行拼图

✅ 仅使用 PySide6（已去掉 PyQt5）
✅ 已移除“贴二维码”相关全部代码
✅ 仅调用 get_best5（不再有“全拼/整拼”说法）
"""

import os
import sys
import time
import shutil
import traceback
import subprocess
import re
from datetime import datetime

# ====== 仅使用 get_best5（必须同目录）======
import get_best5

# ---- 强制本进程输出编码，避免混码导致子进程/父进程互相坑 ----
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_NAME_P1 = "over_test_p1.pdf"
OUT_NAME_P2 = "over_test_p2.pdf"

# -------------------------
# Qt imports (ONLY PySide6)
# -------------------------
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLineEdit, QPushButton, QLabel, QTextEdit, QFileDialog,
    QMessageBox, QProgressBar, QSplitter,
    QFrame, QDialog, QRadioButton, QButtonGroup
)


# -------------------------
# 工具函数
# -------------------------
def ensure_dir(path):
    if path and (not os.path.isdir(path)):
        os.makedirs(path)


def is_output_like_pdf(filename_lower):
    # 排除输出文件，避免重复处理
    if filename_lower.startswith("over_test") and filename_lower.endswith(".pdf"):
        return True
    return False


def list_input_pdfs(folder):
    """列出可作为输入的PDF（排除 over_test*.pdf）"""
    if not os.path.isdir(folder):
        return []
    out = []
    for fn in os.listdir(folder):
        l = fn.lower()
        if not l.endswith(".pdf"):
            continue
        if is_output_like_pdf(l):
            continue
        p = os.path.join(folder, fn)
        if os.path.isfile(p):
            out.append(p)
    out.sort()
    return out


def list_ai_files(folder):
    """列出 folder 下的 .ai（只扫当前层，不递归；需要递归可改 os.walk）"""
    if not os.path.isdir(folder):
        return []
    out = []
    for fn in os.listdir(folder):
        if fn.lower().endswith(".ai"):
            p = os.path.join(folder, fn)
            if os.path.isfile(p):
                out.append(p)
    out.sort()
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
    """给 get_best5 注入路径。"""
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
# ✅ PDF -> AI（仅改后缀名）
# -------------------------
def rename_pdf_to_ai_in_dir(in_dir, recursive=True, overwrite=False, log_cb=None, stop_flag=None):
    """
    只做“把 .pdf/.PDF 改成 .ai”：
    - 不是格式转换，只是改扩展名
    """
    if not os.path.isdir(in_dir):
        if log_cb:
            log_cb("[FATAL] not found: %s\n" % in_dir)
        return {"ok": 0, "skip": 0, "err": 1}

    cnt_ok = cnt_skip = cnt_err = 0

    def iter_pdf_files():
        if recursive:
            for root, _, files in os.walk(in_dir):
                for fn in files:
                    lf = fn.lower()
                    if lf.endswith(".pdf"):
                        yield os.path.join(root, fn)
        else:
            for fn in os.listdir(in_dir):
                lf = fn.lower()
                if lf.endswith(".pdf"):
                    p = os.path.join(in_dir, fn)
                    if os.path.isfile(p):
                        yield p

    for pdf_path in iter_pdf_files():
        if stop_flag and stop_flag():
            break
        try:
            base, _ = os.path.splitext(pdf_path)
            ai_path = base + ".ai"

            if os.path.exists(ai_path):
                if overwrite:
                    os.remove(ai_path)
                else:
                    if log_cb:
                        log_cb("[SKIP exists] %s\n" % ai_path)
                    cnt_skip += 1
                    continue

            os.replace(pdf_path, ai_path)
            if log_cb:
                log_cb("[OK] %s -> %s\n" % (pdf_path, ai_path))
            cnt_ok += 1
        except Exception as e:
            if log_cb:
                log_cb("[ERR] %s  %s\n" % (pdf_path, repr(e)))
            cnt_err += 1

    if log_cb:
        log_cb("[DONE rename] OK=%d SKIP=%d ERR=%d\n" % (cnt_ok, cnt_skip, cnt_err))
    return {"ok": cnt_ok, "skip": cnt_skip, "err": cnt_err}


# -------------------------
# ✅ 运行 JSX/VBS 批处理 AI
# -------------------------
def run_jsx_batch(ai_dir, cx, vbs_path, jsx_path, out_dir, log_cb=None, stop_flag=None, proc_setter=None):
    """
    对 ai_dir 下每个 .ai 执行：
      cscript //nologo run_ai.vbs AItest_ai.jsx "2;AI全路径;输出目录"
    """
    if not os.path.isfile(vbs_path):
        raise RuntimeError("找不到 vbs：%s" % vbs_path)
    if not os.path.isfile(jsx_path):
        raise RuntimeError("找不到 jsx：%s" % jsx_path)

    ai_files = list_ai_files(ai_dir)
    if not ai_files:
        if log_cb:
            log_cb("⚠️ AI目录没有 .ai 文件：%s\n" % ai_dir)
        return 0

    total = len(ai_files)
    ok_count = 0
    fail_count = 0

    for i, ai_path in enumerate(ai_files, start=1):
        if stop_flag and stop_flag():
            break

        data = "%s;%s;%s" % (str(cx), ai_path, out_dir)
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
            if log_cb:
                log_cb("❌ JSX 子进程失败 rc=%s  file=%s\n" % (rc, ai_path))
            # 单文件失败不终止全批次，继续处理下一个
            if log_cb:
                log_cb("⚠️ 跳过失败文件，继续后续文件。\n")
            continue

        if ret_failed:
            fail_count += 1
            if log_cb:
                log_cb("⚠️ JSX 返回失败，已跳过：file=%s  ret=%s\n" % (ai_path, ret_payload))
            continue

        ok_count += 1

        # 让日志里也能显示总体进度（和你原来的 PROGRESS 机制兼容）
        if log_cb:
            log_cb("PROGRESS: %d / %d\n" % (i, total))

    if log_cb:
        log_cb("JSX批处理统计：成功=%d 失败=%d 总数=%d\n" % (ok_count, fail_count, total))

    # 全部失败时返回非0，提醒上层停止
    if total > 0 and ok_count == 0 and fail_count > 0:
        return 5

    return 0


# -------------------------
# CLI 子进程模式：仅运行 get_best5
# -------------------------
def _cli_run_algo(work_dir, out1, out2):
    apply_paths_to_module(get_best5, work_dir, work_dir, out1, out2)
    print("=== RUN get_best5.py (拼图) ===")
    get_best5.main()
    return 0


def _is_cli_mode():
    return ("--run-algo" in sys.argv)


def _cli_entry():
    # 简单解析参数（避免引入 argparse）
    if "--run-algo" in sys.argv:
        def _get(flag, default=""):
            if flag in sys.argv:
                j = sys.argv.index(flag)
                return sys.argv[j + 1] if j + 1 < len(sys.argv) else default
            return default

        work_dir = _get("--work", "")
        out1 = _get("--out1", "")
        out2 = _get("--out2", "")

        if not work_dir or not out1 or not out2:
            print("ERR: missing args for --run-algo")
            return 2

        try:
            return _cli_run_algo(work_dir, out1, out2)
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

    def _transfer_input_pdfs(self, src_dir, dst_dir, mode="copy"):
        ensure_dir(dst_dir)
        files = list_input_pdfs(src_dir)
        moved = []
        for src_path in files:
            if self._stop:
                break
            base = os.path.basename(src_path)
            dst_path = unique_dest_path(dst_dir, base)
            if mode == "move":
                shutil.move(src_path, dst_path)
            else:
                shutil.copy2(src_path, dst_path)
            moved.append(dst_path)
        return moved

    def run(self):
        try:
            cfg = self.cfg
            ai_dir = cfg["ai_dir"]          # ✅ 你新增的AI目录（也是 PDF 输入/输出目录）
            test_root = cfg["test_root"]
            out1 = cfg["out1"]
            out2 = cfg["out2"]
            mode = cfg["mode"]
            cx = cfg.get("cx", 2)

            script_dir = os.path.dirname(os.path.abspath(__file__))
            # 优先使用 run/ 下的稳健版 JSX（对刀线识别与异常容错更好）
            jsx_path = os.path.join(script_dir, "run", "AItest_ai.jsx")
            if not os.path.isfile(jsx_path):
                jsx_path = os.path.join(script_dir, "AItest_ai.jsx")
            vbs_path = os.path.join(script_dir, "run_ai.vbs")

            self.sig_status.emit("Running", "Preparing")
            self.sig_progress.emit(0.0)

            ensure_dir(test_root)
            ensure_dir(out1)
            ensure_dir(out2)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            work_dir = os.path.join(test_root, "work_" + ts)
            ensure_dir(work_dir)

            self.sig_log.emit("=== CONFIG ===\n")
            self.sig_log.emit("AI目录（PDF->AI + JSX导出PDF + 拼图输入PDF）: %s\n" % ai_dir)
            self.sig_log.emit("工作目录 TEST_DIR(work)               : %s\n" % work_dir)
            self.sig_log.emit("输出拼图 DEST_DIR1                   : %s\n" % out1)
            self.sig_log.emit("输出刀线 DEST_DIR2                   : %s\n" % out2)
            self.sig_log.emit("传输模式 TRANSFER_MODE               : %s\n" % mode)
            self.sig_log.emit("JSX                                  : %s\n" % jsx_path)
            self.sig_log.emit("VBS                                  : %s\n" % vbs_path)
            self.sig_log.emit("CX                                   : %s\n" % str(cx))
            self.sig_log.emit("运行算法 ALGO                         : get_best5 (拼图)\n")
            self.sig_log.emit("================\n\n")

            if not os.path.isdir(ai_dir):
                self.sig_log.emit("❌ AI目录不存在：%s\n" % ai_dir)
                self.sig_done.emit(2)
                return

            # 1) ✅ PDF -> AI（改后缀）
            self.sig_status.emit("Running", "PDF->AI")
            self.sig_log.emit("=== STEP1: RENAME .PDF -> .AI (just extension) ===\n")
            rename_pdf_to_ai_in_dir(
                ai_dir, recursive=True, overwrite=False,
                log_cb=lambda s: self.sig_log.emit(s),
                stop_flag=lambda: self._stop
            )
            if self._stop:
                self.sig_log.emit("⛔ 已停止：不再继续。\n")
                self.sig_done.emit(0)
                return

            # 2) ✅ 执行 JSX 批处理（导出 PDF 到同一目录）
            self.sig_status.emit("Running", "AI->PDF (JSX)")
            self.sig_log.emit("\n=== STEP2: RUN JSX (AI -> PDF export) ===\n")
            rc_jsx = run_jsx_batch(
                ai_dir=ai_dir,
                cx=cx,
                vbs_path=vbs_path,
                jsx_path=jsx_path,
                out_dir=ai_dir,  # ✅ 输出PDF目录=AI目录（你要求：PDF输入目录就是 JSX 输出目录）
                log_cb=lambda s: self.sig_log.emit(s),
                stop_flag=lambda: self._stop,
                proc_setter=self._set_current_proc
            )
            if self._stop:
                self.sig_log.emit("⛔ 已停止：JSX阶段中断。\n")
                self.sig_done.emit(0)
                return
            if rc_jsx != 0:
                self.sig_log.emit("❌ JSX 执行失败 rc=%s\n" % rc_jsx)
                self.sig_done.emit(rc_jsx)
                return

            # 3) ✅ 检查 JSX 导出的 PDF
            pdfs_now = list_input_pdfs(ai_dir)
            if not pdfs_now:
                self.sig_log.emit("\n⚠️ JSX运行后未发现可处理PDF（排除 over_test*.pdf）：%s\n" % ai_dir)
                self.sig_done.emit(0)
                return

            # 4) Transfer PDF 到 work
            self.sig_status.emit("Running", "Transfer")
            self.sig_log.emit("\n=== STEP3: TRANSFER INPUT PDFS ===\n")
            moved = self._transfer_input_pdfs(ai_dir, work_dir, mode=("move" if mode == "move" else "copy"))
            self.sig_log.emit("传输完成：%d 个PDF\n\n" % len(moved))
            if self._stop:
                self.sig_log.emit("⛔ 已停止：不再继续运行算法。\n")
                self.sig_done.emit(0)
                return

            # 5) Algo subprocess
            self.sig_status.emit("Running", "Algorithm")
            self.sig_log.emit("=== STEP4: RUN ALGO (SUBPROCESS) ===\n")
            cmd_algo = [
                sys.executable, "-u", os.path.abspath(__file__),
                "--run-algo",
                "--work", work_dir,
                "--out1", out1,
                "--out2", out2
            ]
            rc = self._spawn_and_stream(cmd_algo)
            if self._stop:
                self.sig_log.emit("⛔ 已停止：算法阶段中断。\n")
                self.sig_done.emit(0)
                return
            if rc != 0:
                self.sig_log.emit("\n❌ 算法子进程失败 (rc=%s)\n" % rc)
                self.sig_done.emit(rc)
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
        self.setWindowTitle("说明：AI预处理 + 拼图运行器")
        self.resize(900, 560)

        layout = QVBoxLayout(self)
        self.txt = QTextEdit(self)
        self.txt.setReadOnly(True)
        self.txt.setFont(QFont("Consolas", 11))
        layout.addWidget(self.txt)

        btn = QPushButton("返回", self)
        btn.clicked.connect(self.accept)
        layout.addWidget(btn, alignment=Qt.AlignRight)

        help_text = (
            "【运行顺序（严格按此顺序）】\n"
            "1）选择“输入AI文件夹”（也是 JSX 导出 PDF 的目录，也是拼图输入 PDF 的目录）\n"
            "2）自动把该目录里所有 .pdf/.PDF 改后缀为 .ai（仅改扩展名，不是格式转换）\n"
            "3）对每个 .ai 执行：cscript //nologo run_ai.vbs AItest_ai.jsx \"2;AI全路径;输出目录\"\n"
            "4）把该目录里导出的 PDF（排除 over_test*.pdf）复制/移动到 work 目录\n"
            "5）子进程运行 get_best5 进行拼图输出\n\n"
            "【注意】\n"
            "- run.py、AItest_ai.jsx、run_ai.vbs 必须在同一目录\n"
            "- 若 JSX 需要动态 OUT_DIR，请按我给你的“JSX最小改法”改一行\n"
            "- 进度条：JSX阶段会输出 PROGRESS: a / b；算法阶段若 get_best5 也输出 PROGRESS，会继续更新\n"
        )
        self.txt.setPlainText(help_text)


# -------------------------
# Main Window
# -------------------------
class RollWidget(QWidget):
    def __init__(self):
        super(RollWidget, self).__init__()

        self.setWindowTitle("印客链 - AI预处理 + 拼图运行器")

        ico_path = r"C:\Users\wzqy\PycharmProjects\inklink\icon.ico"
        if os.path.isfile(ico_path):
            try:
                self.setWindowIcon(QIcon(ico_path))
            except Exception:
                pass

        self.resize(1100, 780)
        self.worker = None

        self._build_ui()
        self._apply_qss()

        # 默认值
        self.ed_ai.setText(r"D:\test_data\iest")
        self.ed_dest.setText(self.ed_ai.text().strip())
        self.ed_test.setText(r"D:\test_data\test")
        self.ed_out1.setText(r"D:\test_data\gest")
        self.ed_out2.setText(r"D:\test_data\pest")
        self.rb_copy.setChecked(True)

        self._set_status("Idle", "Ready")
        self._set_progress_indeterminate(False)
        self.clear_log()

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
        lb_title = QLabel("印客链 - AI预处理 + 拼图运行器")
        f = QFont("Microsoft YaHei UI", 18)
        f.setBold(True)
        lb_title.setFont(f)
        lb_sub = QLabel("顺序：PDF->AI（改后缀）→ JSX导出PDF → 子进程拼图 get_best5")
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

        # ✅ 新增：AI目录（也是 PDF 输入/输出目录）
        self.ed_ai = add_path_row(0, "输入AI目录（=JSX输出PDF目录=拼图输入PDF目录）")
        self.ed_ai.textChanged.connect(self._sync_dest_to_ai)

        # PDF输入目录：自动同步显示（只读）
        self.ed_dest = add_path_row(1, "输入PDF目录（自动=AI目录）", pick=False)
        self.ed_dest.setReadOnly(True)

        self.ed_test = add_path_row(2, "工作根目录")
        self.ed_out1 = add_path_row(3, "输出拼图目录")
        self.ed_out2 = add_path_row(4, "输出刀线目录")

        grid.setColumnStretch(1, 1)
        left_layout.addWidget(gp_paths)

        gp_opts = QGroupBox("运行选项")
        vopts = QVBoxLayout(gp_opts)

        row_mode = QHBoxLayout()
        row_mode.addWidget(QLabel("传输模式："))
        self.rb_copy = QRadioButton("复制 copy")
        self.rb_move = QRadioButton("移动 move")
        grp_mode = QButtonGroup(self)
        grp_mode.addButton(self.rb_copy, 0)
        grp_mode.addButton(self.rb_move, 1)
        row_mode.addWidget(self.rb_copy)
        row_mode.addWidget(self.rb_move)
        row_mode.addStretch(1)
        vopts.addLayout(row_mode)

        left_layout.addWidget(gp_opts)

        gp_btn = QGroupBox("操作")
        hb = QHBoxLayout(gp_btn)

        self.btn_start = QPushButton("▶ 开始运行")
        self.btn_start.setObjectName("btnStart")
        self.btn_start.clicked.connect(self.start)

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

        tip = QLabel("提示：AI目录会自动作为PDF输入/输出目录；JSX/VBS需与run.py同目录。")
        tip.setStyleSheet("color:#94a3b8;")
        left_layout.addWidget(tip)
        left_layout.addStretch(1)

        splitter.addWidget(left)

        # Right panel (log)
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        gp_log = QGroupBox("运行日志（子进程 stdout 实时回传）")
        vlog = QVBoxLayout(gp_log)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Consolas", 11))
        vlog.addWidget(self.log)

        right_layout.addWidget(gp_log)
        splitter.addWidget(right)

        splitter.setSizes([440, 660])

    def _sync_dest_to_ai(self):
        self.ed_dest.setText(self.ed_ai.text().strip())

    def _pick_dir(self, line_edit: QLineEdit):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹", line_edit.text().strip() or os.getcwd())
        if d:
            line_edit.setText(d)

    def _set_status(self, status, phase):
        self.lb_status.setText(status)
        self.lb_phase.setText(phase)

    def _set_progress_indeterminate(self, on: bool):
        if on:
            self.pb.setRange(0, 0)   # busy
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

    def _get_mode_value(self):
        return "move" if self.rb_move.isChecked() else "copy"

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
        elif ("⚠️" in line) or line.startswith("WARN:") or line.startswith("[SKIP"):
            color = "#fbbf24"
        elif ("✅" in line) or ("OK" in line) or line.startswith("[OK]"):
            color = "#34d399"
        elif "⛔" in line:
            color = "#fbbf24"

        # 阶段推断（用于进度条 busy）
        if "STEP1:" in line:
            self._set_status("Running", "PDF->AI")
            self._set_progress_indeterminate(True)
        elif "STEP2:" in line:
            self._set_status("Running", "AI->PDF (JSX)")
            self._set_progress_indeterminate(True)
        elif "STEP3:" in line or "=== STEP3" in line:
            self._set_status("Running", "Transfer")
            self._set_progress_indeterminate(True)
        elif "STEP4:" in line or "=== STEP4" in line:
            self._set_status("Running", "Algorithm")
            self._set_progress_indeterminate(True)

        self.log.append('<span style="color:%s">%s</span>' % (color, self._html_escape(line)))

    def start(self):
        if self.worker is not None:
            return

        ai_dir = self.ed_ai.text().strip()
        test_root = self.ed_test.text().strip()
        out1 = self.ed_out1.text().strip()
        out2 = self.ed_out2.text().strip()
        mode = self._get_mode_value()

        if not os.path.isdir(ai_dir):
            QMessageBox.critical(self, "错误", "输入AI目录不存在：\n" + ai_dir)
            return
        if not test_root:
            QMessageBox.critical(self, "错误", "工作根目录不能为空")
            return

        ensure_dir(test_root)
        ensure_dir(out1)
        ensure_dir(out2)

        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self._set_status("Running", "Preparing")
        self._set_progress_indeterminate(True)
        self.pb.setValue(0)

        self.clear_log()

        cfg = {
            "ai_dir": ai_dir,
            "test_root": test_root,
            "out1": out1,
            "out2": out2,
            "mode": mode,
            "cx": 2,   # 你给的 data 前缀就是 2；需要改就改这里
        }

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

        self.btn_start.setEnabled(True)
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
        if self.worker is not None:
            ret = QMessageBox.question(
                self, "退出", "正在运行中，确定要退出吗？（会终止子进程）",
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