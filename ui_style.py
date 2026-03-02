# -*- coding: utf-8 -*-
"""
UI 样式
"""

def build_qss():
    return """
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


def apply_qss(widget):
    widget.setStyleSheet(build_qss())