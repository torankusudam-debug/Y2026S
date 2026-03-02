# -*- coding: utf-8 -*-
"""
说明弹窗
"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton


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
            "- 若 JSX 需要动态 OUT_DIR，请按你之前那段“JSX最小改法”让 OUT_DIR = 第三段参数\n"
            "- 进度条：JSX阶段会输出 PROGRESS: a / b；算法阶段若 get_best5 也输出 PROGRESS，会继续更新\n"
        )
        self.txt.setPlainText(help_text)