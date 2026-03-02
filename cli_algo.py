# -*- coding: utf-8 -*-
"""
CLI 子进程模式：只运行拼图算法 get_best5
用法（内部调用）：
  python run.py --run-algo --work <dir> --out1 <dir> --out2 <dir>
"""

import os
import traceback

import get_best5
from utils_core import apply_paths_to_module


def _get_arg(argv, flag, default=""):
    if flag in argv:
        j = argv.index(flag)
        return argv[j + 1] if j + 1 < len(argv) else default
    return default


def _cli_run_algo(work_dir, out1, out2):
    apply_paths_to_module(get_best5, work_dir, work_dir, out1, out2)
    print("=== RUN get_best5.py (拼图) ===")
    get_best5.main()
    return 0


def cli_entry(argv):
    work_dir = _get_arg(argv, "--work", "")
    out1 = _get_arg(argv, "--out1", "")
    out2 = _get_arg(argv, "--out2", "")

    if not work_dir or not out1 or not out2:
        print("ERR: missing args for --run-algo")
        return 2

    try:
        return _cli_run_algo(work_dir, out1, out2)
    except Exception:
        print("❌ ALGO EXCEPTION:\n" + traceback.format_exc())
        return 1