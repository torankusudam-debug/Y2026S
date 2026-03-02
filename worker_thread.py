# -*- coding: utf-8 -*-
"""
后台工作线程：按顺序做
1) PDF->AI 改后缀
2) 批量跑 JSX 导出 PDF（输出目录=AI目录）
3) 把 PDF 复制/移动到 work
4) 子进程运行 run.py --run-algo (get_best5)
"""

import os
import sys
import shutil
import traceback
import subprocess
import re
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from utils_core import ensure_dir, list_input_pdfs, unique_dest_path
from ai_preprocess import rename_pdf_to_ai_in_dir, run_jsx_batch


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
            ai_dir = cfg["ai_dir"]
            test_root = cfg["test_root"]
            out1 = cfg["out1"]
            out2 = cfg["out2"]
            mode = cfg["mode"]
            cx = cfg.get("cx", 2)

            script_dir = os.path.dirname(os.path.abspath(__file__))
            jsx_path = os.path.join(script_dir, "AItest_ai.jsx")
            vbs_path = os.path.join(script_dir, "run_ai.vbs")
            entry_run_py = os.path.join(script_dir, "run.py")

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

            # 1) PDF -> AI（改后缀）
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

            # 2) 执行 JSX 批处理（导出 PDF 到同一目录）
            self.sig_status.emit("Running", "AI->PDF (JSX)")
            self.sig_log.emit("\n=== STEP2: RUN JSX (AI -> PDF export) ===\n")
            rc_jsx = run_jsx_batch(
                ai_dir=ai_dir,
                cx=cx,
                vbs_path=vbs_path,
                jsx_path=jsx_path,
                out_dir=ai_dir,  # 输出PDF目录=AI目录
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

            # 3) 检查 JSX 导出的 PDF
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

            # 5) Algo subprocess（调用入口 run.py --run-algo）
            self.sig_status.emit("Running", "Algorithm")
            self.sig_log.emit("=== STEP4: RUN ALGO (SUBPROCESS) ===\n")
            cmd_algo = [
                sys.executable, "-u", entry_run_py,
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