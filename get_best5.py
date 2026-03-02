# -*- coding: utf-8 -*-
"""
get_best5.py（整拼优先 + 整拼失败自动全拼）

✅ 输出规则（保持不变）：
1) 全拼：所有“整拼无解”的内容，汇总输出到同一个 PDF
   - 输出文件名固定为：全拼.pdf
   - pdf1（图形页/实图）与 pdf2（轮廓页）各输出一份

2) 整拼：每个“源PDF”单独输出
   - 输出文件名：与源文件名一致（仅扩展名改为 .pdf）
   - pdf1 与 pdf2 分别输出到各自目录
   - 若某个源PDF里有多个页对 (0,1)(2,3)...：该源PDF内所有“整拼可行”的页对会连续写进同一个输出PDF里
"""

import gb5_config as C
from gb5_runner import run as _run
from gb5_runner import main as _main


def _sync_globals():
    # 兼容：外部若读取 get_best5.DEST_DIR 等，尽量保持同步
    g = globals()
    g["DEST_DIR"] = C.DEST_DIR
    g["IN_PDF_ARCHIVE_DIR"] = C.IN_PDF_ARCHIVE_DIR
    g["DEST_DIR1"] = C.DEST_DIR1
    g["DEST_DIR2"] = C.DEST_DIR2
    g["OUT_PDF_P1"] = C.OUT_PDF_P1
    g["OUT_PDF_P2"] = C.OUT_PDF_P2


def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    C.set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2)
    _sync_globals()


def run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None):
    _sync_globals()
    return _run(cfg=cfg, input_pdfs=input_pdfs, progress_cb=progress_cb, log_cb=log_cb)


def main():
    _sync_globals()
    return _main()


# 初始化一次
_sync_globals()

if __name__ == "__main__":
    main()