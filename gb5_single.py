# -*- coding: utf-8 -*-
"""
gb5_single.py —— 整拼模块

整拼规则：
  - 同一个源文件的图形排入块（≤320×464mm）
  - 块间间隔 BLOCK_GAP (10mm)
  - 每个块的右上角放二维码，左上角放单号
  - 刀线只标注在左侧边缘和顶部边缘，长度5mm
  - 纸宽固定600mm，就算没排满一排也保持600mm
  - N≥10 时才进入整拼
"""
import math
import fitz

import gb5_config as C
from gb5_utils import mm_to_pt, pick_cjk_fontfile, trim_text_to_width


# ============================================================
# 块内排列：在一个 BLOCK_W × BLOCK_H 的块内能放多少个图形
# ============================================================

def _block_layout(z_mm, h_mm):
    """
    计算一个块内能放多少个图形。
    z_mm: 图形宽(mm)
    h_mm: 图形高(mm)
    返回: (cols_per_block, rows_per_block, block_actual_w, block_actual_h)
    """
    if z_mm <= 0 or h_mm <= 0:
        return 0, 0, 0, 0
    cols = int(C.BLOCK_W // z_mm)
    rows = int(C.BLOCK_H // h_mm)
    if cols <= 0 or rows <= 0:
        return 0, 0, 0, 0
    actual_w = cols * z_mm
    actual_h = rows * h_mm
    return cols, rows, actual_w, actual_h


def _blocks_per_row(block_actual_w, paper_w):
    """
    一排（纸宽 paper_w）能放多少个块。
    块间间隔 BLOCK_GAP。
    """
    if block_actual_w <= 0:
        return 0
    bpr = 1
    while (bpr + 1) * block_actual_w + bpr * C.BLOCK_GAP <= paper_w + 0.01:
        bpr += 1
    return bpr


def solve_single_type_no_waste(outer_w, outer_h, need_count):
    """
    整拼求解：找到最佳排列方案。

    尝试两种方向（不旋转/旋转90度），选择总面积最小的方案。

    返回 best dict 或 None。
    """
    best = None

    for ori in (0, 1):
        z = outer_w if ori == 0 else outer_h
        h = outer_h if ori == 0 else outer_w

        cols_blk, rows_blk, blk_w, blk_h = _block_layout(z, h)
        if cols_blk <= 0 or rows_blk <= 0:
            continue

        cap_per_block = cols_blk * rows_blk
        bpr = _blocks_per_row(blk_w, C.SINGLE_W_MAX)
        if bpr <= 0:
            continue

        # 一排的宽度
        row_w = bpr * blk_w + (bpr - 1) * C.BLOCK_GAP
        # 纸宽固定600mm
        W = C.SINGLE_W_MAX

        # 计算需要多少排（垂直方向）
        blocks_needed = int(math.ceil(need_count / float(cap_per_block)))

        # 一张纸上能放多少排块
        max_block_rows = 1
        while True:
            test_h = max_block_rows * blk_h + (max_block_rows - 1) * C.BLOCK_GAP
            if test_h > C.SINGLE_H_MAX + 0.01:
                max_block_rows -= 1
                break
            if max_block_rows * bpr >= blocks_needed:
                break
            max_block_rows += 1

        if max_block_rows <= 0:
            max_block_rows = 1

        blocks_per_page = max_block_rows * bpr
        cap_per_page = blocks_per_page * cap_per_block
        pages = int(math.ceil(need_count / float(cap_per_page)))

        # 实际高度
        H = max_block_rows * blk_h + (max_block_rows - 1) * C.BLOCK_GAP
        if H > C.SINGLE_H_MAX + 0.01:
            continue

        area = pages * W * H

        cand = {
            "ori": ori,
            "z": float(z),
            "h": float(h),
            "cols_blk": cols_blk,
            "rows_blk": rows_blk,
            "blk_w": float(blk_w),
            "blk_h": float(blk_h),
            "bpr": bpr,              # 每排块数
            "block_rows": max_block_rows,  # 垂直方向块排数
            "cap_per_block": cap_per_block,
            "cap_per_page": cap_per_page,
            "W": float(W),
            "H": float(H),
            "pages": pages,
            "total_area": float(area),
        }

        if best is None or cand["total_area"] < best["total_area"] - 1e-9:
            best = cand
        elif abs(cand["total_area"] - best["total_area"]) <= 1e-9:
            if cand["H"] < best["H"] - 1e-9:
                best = cand

    return best


def build_single_placements(type_id, best):
    """
    构建整拼的放置列表。

    返回:
      placements: list of dict，每个图形的位置
      blocks: list of dict，每个块的位置和尺寸（用于标注二维码和单号）
    """
    z = best["z"]
    h = best["h"]
    cols_blk = best["cols_blk"]
    rows_blk = best["rows_blk"]
    blk_w = best["blk_w"]
    blk_h = best["blk_h"]
    bpr = best["bpr"]
    block_rows = best["block_rows"]

    placements = []
    blocks = []

    for br in range(block_rows):
        block_y = br * (blk_h + C.BLOCK_GAP)
        for bc in range(bpr):
            block_x = bc * (blk_w + C.BLOCK_GAP)

            blocks.append({
                "x": block_x,
                "y": block_y,
                "w": blk_w,
                "h": blk_h,
            })

            # 块内排列图形
            for row in range(rows_blk):
                for col in range(cols_blk):
                    px = block_x + col * z
                    py = block_y + row * h
                    placements.append({
                        "type": type_id,
                        "x": px,
                        "y": py,
                        "w": z,
                        "h": h,
                        "rot": (best["ori"] == 1),
                        "block_idx": len(blocks) - 1,
                    })

    return placements, blocks


def compute_edges_from_placements(placements):
    """
    从图片列表计算刀线边界，保证相邻图片可分开。
    """
    xs = set()
    ys = set()
    for p in placements:
        xs.add(round(p["x"], 2))
        xs.add(round(p["x"] + p["w"], 2))
        ys.add(round(p["y"], 2))
        ys.add(round(p["y"] + p["h"], 2))
    return sorted(xs), sorted(ys)


def draw_label_in_block(page, text, block_x, block_y, block_w):
    """在块的左上角绘制单号标签。"""
    fontfile = pick_cjk_fontfile()
    xpt = mm_to_pt(block_x + 1.0)
    ypt = mm_to_pt(block_y + C.QR_H - 1.5)
    max_w = mm_to_pt(block_w - C.QR_W - 2.0)
    txt = trim_text_to_width(text, C.LABEL_FONT_SIZE, max_w)

    if fontfile:
        page.insert_text(fitz.Point(xpt, ypt), txt,
                         fontsize=C.LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt,
                         fontsize=C.LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))


def append_pages_single(out_doc, best, placements, blocks, pages, qr_bytes, label_text, img_bytes):
    """
    整拼绘制页面。

    标注规则：
      - 每个块的右上角放二维码
      - 每个块的左上角放单号
      - 刀线只标注在左侧边缘和顶部边缘，长度5mm
    """
    W = best["W"]
    H = best["H"]
    Wpt = mm_to_pt(W)
    Hpt = mm_to_pt(H)

    x_edges, y_edges = compute_edges_from_placements(placements)
    ml = mm_to_pt(C.MARK_LEN)  # 刀线长度 5mm

    for _ in range(pages):
        page = out_doc.new_page(width=Wpt, height=Hpt)

        # 纸张外框
        page.draw_rect(fitz.Rect(0, 0, Wpt, Hpt), color=(0, 0, 0), width=1.0)

        # 绘制所有图形
        for p in placements:
            x = p["x"]; y = p["y"]; w = p["w"]; h = p["h"]
            outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))

            if C.DRAW_PART_OUTER_BOX:
                page.draw_rect(outer, color=(0, 0, 0), width=0.5)

            inner = fitz.Rect(
                outer.x0 + mm_to_pt(C.INNER_MARGIN_MM),
                outer.y0 + mm_to_pt(C.INNER_MARGIN_MM),
                outer.x1 - mm_to_pt(C.INNER_MARGIN_MM),
                outer.y1 - mm_to_pt(C.INNER_MARGIN_MM),
            )

            if img_bytes:
                page.insert_image(inner, stream=img_bytes, keep_proportion=True,
                                  rotate=(90 if p["rot"] else 0))

        # 每个块标注二维码（右上角）和单号（左上角）
        for b in blocks:
            bx = b["x"]
            by = b["y"]
            bw = b["w"]

            # 二维码放块的右上角
            qr_rect = fitz.Rect(
                mm_to_pt(bx + bw - C.QR_W),
                mm_to_pt(by),
                mm_to_pt(bx + bw),
                mm_to_pt(by + C.QR_H),
            )
            page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
            if qr_bytes:
                page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

            # 单号放块的左上角
            if label_text:
                draw_label_in_block(page, label_text, bx, by, bw)

        # ---- 刀线：只在顶部边缘和左侧边缘，长度5mm ----
        # 顶部边缘：垂直刀线
        for xx in x_edges:
            xpt = mm_to_pt(xx)
            if 0 <= xpt <= Wpt:
                page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, ml),
                               color=(0, 0, 0), width=1.0)

        # 左侧边缘：水平刀线
        for yy in y_edges:
            ypt = mm_to_pt(yy)
            if 0 <= ypt <= Hpt:
                page.draw_line(fitz.Point(0, ypt), fitz.Point(ml, ypt),
                               color=(0, 0, 0), width=1.0)
