# -*- coding: utf-8 -*-
"""
AI 预处理相关：
1) PDF->AI（仅改后缀名）
2) 批量运行 JSX（cscript run_ai.vbs ...）
"""

import os
import subprocess

from utils_core import list_ai_files


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
                    if fn.lower().endswith(".pdf"):
                        yield os.path.join(root, fn)
        else:
            for fn in os.listdir(in_dir):
                if fn.lower().endswith(".pdf"):
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
        if proc_setter:
            proc_setter(p)

        try:
            for line in p.stdout:
                if log_cb:
                    log_cb(line)
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
            if log_cb:
                log_cb("❌ JSX 子进程失败 rc=%s  file=%s\n" % (rc, ai_path))
            return rc

        # 让日志里也能显示总体进度（和原 PROGRESS 机制兼容）
        if log_cb:
            log_cb("PROGRESS: %d / %d\n" % (i, total))

    return 0