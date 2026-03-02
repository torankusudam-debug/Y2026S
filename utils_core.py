# -*- coding: utf-8 -*-
"""
通用工具函数：路径、列表、注入配置等
"""


import time
import shutil
import os

def ensure_dir(p):
    """确保目录存在（p 为空则忽略）"""
    if not p:
        return
    if not os.path.isdir(p):
        os.makedirs(p)

OUT_NAME_P1 = "over_test_p1.pdf"
OUT_NAME_P2 = "over_test_p2.pdf"


def ensure_dir(path):
    if path and (not os.path.isdir(path)):
        os.makedirs(path)


def is_output_like_pdf(filename_lower):
    # 排除输出文件，避免重复处理
    return filename_lower.startswith("over_test") and filename_lower.endswith(".pdf")


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
    """列出 folder 下的 .ai（只扫当前层）"""
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
    """给 get_best5 注入路径（保持原 run.py 行为）"""
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