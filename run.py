# -*- coding: utf-8 -*-
"""
入口文件：
- GUI 模式：启动 PySide6 界面
- CLI 子进程模式：--run-algo 运行 get_best5
"""

import sys

# ---- 强制本进程输出编码，避免混码导致子进程/父进程互相坑 ----
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def main():
    # ✅ 子进程 CLI 模式（给 RunnerThread 调用）
    if "--run-algo" in sys.argv:
        from cli_algo import cli_entry
        code = cli_entry(sys.argv)
        raise SystemExit(int(code))

    # ✅ GUI 模式
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QFont
    from ui_window import RollWidget

    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei UI", 11))

    w = RollWidget()
    w.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()