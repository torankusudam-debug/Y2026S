# -*- coding: utf-8 -*-
"""
主窗口 UI（保持原功能不变）
"""

import os
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLineEdit, QPushButton, QLabel, QTextEdit, QFileDialog,
    QMessageBox, QProgressBar, QSplitter,
    QFrame, QDialog, QRadioButton, QButtonGroup
)

from worker_thread import RunnerThread
from utils_core import ensure_dir
from ui_help import HelpDialog
from ui_style import apply_qss


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
        apply_qss(self)

        # 默认值（保持原来一致）
        self.ed_ai.setText(r"D:\test_data\iest")
        self.ed_test.setText(r"D:\test_data\test")

        self.ed_out1.setText(r"D:\test_data\gest")
        self.ed_out2.setText(r"D:\test_data\pest")
        self.rb_copy.setChecked(True)

        self._set_status("Idle", "Ready")
        self._set_progress_indeterminate(False)
        self.clear_log()

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

        # ✅ AI目录（也是 PDF 输入/输出目录）
        # ✅ AI目录（也是 PDF 输入/输出目录）
        self.ed_ai = add_path_row(0, "PDF输入目录")

        self.ed_test = add_path_row(1, "已处理文件目录")
        self.ed_out1 = add_path_row(2, "输出拼图目录")
        self.ed_out2 = add_path_row(3, "输出刀线目录")

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
        elif ("❌" in line) or line.startswith("ERR:") or ("EXCEPTION" in line) or line.startswith("[FATAL]") or line.startswith("[ERR]"):
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
            "cx": 2,   # 你原来就是传 2；需要改就改这里
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