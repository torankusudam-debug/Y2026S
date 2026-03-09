# -*- coding: utf-8 -*-
"""
gb5_full.py —— 全拼（混拼）模块

核心逻辑：
  - 不同源文件的图形混合排列
  - 图形先填入块（≤320×464mm），块间间隔 BLOCK_GAP (10mm)
  - 块按网格排列：同一列的块宽度相同，同一行的块高度相同
  - 通栏切割：先垂直切（从左到右），再水平切（从上到下）
  - 每一刀都是贯穿整张纸的，不会误伤图形
  - 标注：每个块的单号放右上角，二维码放左上角
  - 刀线只标注在左侧边缘和顶部边缘，长度5mm
"""

import math
import fitz
import gb5_config as C
from gb5_utils import mm_to_pt, _log, pick_cjk_fontfile, trim_text_to_width
from gb5_qr import make_qr_png_bytes


def _pick_piece_spec_for_full(ow, oh):
    """
    为全拼选择单个图形的放置方向，目标是尽量横向塞满。
    返回: (w, h, rot) 或 None
    """
    cands = []
    for rot, w, h in ((False, float(ow), float(oh)), (True, float(oh), float(ow))):
        if w <= C.FULL_W_MAX + 0.01 and h <= C.FULL_H_MAX + 0.01:
            cols = max(1, int(C.FULL_W_MAX // w))
            cands.append((cols, -h, -w, w, h, rot))
    if not cands:
        return None
    cands.sort(reverse=True)
    _, _, _, w, h, rot = cands[0]
    return w, h, rot


def _block_layout_for_entry(ow, oh):
    """
    计算一个块内能放多少个指定尺寸的图形。
    返回: (cols, rows, block_w, block_h)
    """
    if ow <= 0 or oh <= 0:
        return 0, 0, 0, 0
    cols = int(C.BLOCK_W // ow)
    rows = int(C.BLOCK_H // oh)
    if cols <= 0 or rows <= 0:
        return 0, 0, 0, 0
    return cols, rows, cols * ow, rows * oh


def _block_layout_for_entry_rotated(ow, oh):
    """旋转90度后的块内排列。"""
    return _block_layout_for_entry(oh, ow)


def _make_block_for_entry(entry, remaining_count):
    """
    为一个条目创建一个块。

    尝试不旋转和旋转两种方式，选择能放更多图形的方式。
    返回: block_info dict 或 None
    """
    ow = float(entry["ow"])
    oh = float(entry["oh"])
    tid = entry["tid"]

    best_block = None

    for rotated in (False, True):
        if rotated:
            cols, rows, bw, bh = _block_layout_for_entry_rotated(ow, oh)
            img_w, img_h = oh, ow
        else:
            cols, rows, bw, bh = _block_layout_for_entry(ow, oh)
            img_w, img_h = ow, oh

        if cols <= 0 or rows <= 0:
            continue

        cap = cols * rows
        actual_placed = min(cap, remaining_count)

        block = {
            "tid": tid,
            "rotated": rotated,
            "cols": cols,
            "rows": rows,
            "block_w": bw,
            "block_h": bh,
            "img_w": img_w,
            "img_h": img_h,
            "capacity": cap,
            "placed": actual_placed,
        }

        if best_block is None or actual_placed > best_block["placed"]:
            best_block = block
        elif actual_placed == best_block["placed"] and bw * bh < best_block["block_w"] * best_block["block_h"]:
            best_block = block

    return best_block


def pack_full_grid_sheets(entries, log_cb=None):
    """
    将全拼条目按块+网格排列打包成多张sheet。

    排列策略：
      1. 为每个条目生成块（≤320×464mm）
      2. 块按行排列到纸上（纸宽600mm）
      3. 一排中的块高度取最大值（保证通栏水平切割）
      4. 一排中的块从左到右排列，块间间隔BLOCK_GAP
      5. 排与排之间间隔BLOCK_GAP
      6. 高度累加不超过 FULL_H_MAX

    返回:
      sheets: list of dict
      by_tid: dict {tid: entry}
    """
    entries = list(entries)
    by_tid = {}
    for e in entries:
        by_tid[e["tid"]] = e

    # 全局剩余数量
    remaining_global = {}
    for e in entries:
        remaining_global[e["tid"]] = int(e["need"])

    piece_spec = {}
    for e in entries:
        piece_spec[e["tid"]] = _pick_piece_spec_for_full(e["ow"], e["oh"])

    sheets = []

    while any(remaining_global[e["tid"]] > 0 for e in entries):
        # 本张sheet的图形与刀线边界
        sheet_blocks = []
        sheet_placements = []
        col_edges = set()
        row_edges = set()
        placed_count_by_tid = {}

        x_cursor = 0.0
        y_cursor = 0.0
        row_h = 0.0
        any_placed_this_sheet = False
        sheet_full = False

        for e in entries:
            tid = e["tid"]
            spec = piece_spec.get(tid)

            if spec is None:
                if remaining_global.get(tid, 0) > 0:
                    _log(log_cb, "⚠️ 全拼跳过超限尺寸: %s (ow=%.1f oh=%.1f)" %
                         (tid, float(e["ow"]), float(e["oh"])))
                    remaining_global[tid] = 0
                continue

            pw, ph, prot = spec
            while remaining_global.get(tid, 0) > 0:
                # 横向优先：先从左到右拼
                if x_cursor + pw <= C.FULL_W_MAX + 0.01 and y_cursor + max(row_h, ph) <= C.FULL_H_MAX + 0.01:
                    px = x_cursor
                    py = y_cursor
                    sheet_placements.append({
                        "tid": tid,
                        "x": px,
                        "y": py,
                        "w": pw,
                        "h": ph,
                        "rot": prot,
                        "block_idx": -1,
                    })
                    col_edges.add(round(px, 2))
                    col_edges.add(round(px + pw, 2))
                    row_edges.add(round(py, 2))
                    row_edges.add(round(py + ph, 2))
                    x_cursor += pw
                    row_h = max(row_h, ph)
                    remaining_global[tid] -= 1
                    placed_count_by_tid[tid] = placed_count_by_tid.get(tid, 0) + 1
                    any_placed_this_sheet = True
                    continue

                # 当前行放不下就换行；高度不够则结束本张纸
                next_y = y_cursor + row_h
                if row_h <= 0 or next_y + ph > C.FULL_H_MAX + 0.01:
                    sheet_full = True
                    break
                x_cursor = 0.0
                y_cursor = next_y
                row_h = 0.0

            if sheet_full:
                break

        if not any_placed_this_sheet:
            problem = [e["tid"] for e in entries if remaining_global.get(e["tid"], 0) > 0]
            if problem:
                _log(log_cb, "⚠️ 全拼跳过无法放入图形：%s" % str(problem))
                for tid in problem:
                    remaining_global[tid] = 0
            break

        # 计算实际使用的宽高
        W_used = C.FULL_W_MAX
        H_used = y_cursor + row_h

        owner_tid = max(placed_count_by_tid, key=placed_count_by_tid.get) if placed_count_by_tid else ""
        owner_qr = by_tid.get(owner_tid, {}).get("qr_text", "MIX")

        sheet = {
            "owner_tid": owner_tid,
            "qr_text": owner_qr,
            "W": int(W_used),
            "H": int(min(C.FULL_H_MAX, H_used)),
            "placements": sheet_placements,
            "blocks": sheet_blocks,
            "col_edges": sorted(col_edges),
            "row_edges": sorted(row_edges),
        }
        sheets.append(sheet)
        _log(log_cb, "🧩 紧凑全拼出一张：owner=%s W=%d H=%d placed=%d" %
             (owner_tid, sheet["W"], sheet["H"], len(sheet_placements)))

    return sheets, by_tid


def append_one_full_sheet(out_doc, sheet, by_tid, is_contour=False):
    """
    将一个全拼sheet绘制到out_doc中（新增一页）。

    标注规则：
      - 每个块的单号放右上角
      - 每个块的二维码放左上角
      - 刀线只标注在左侧边缘和顶部边缘，长度5mm
    """
    W = int(sheet["W"])
    H = int(sheet["H"])
    Wpt = mm_to_pt(W)
    Hpt = mm_to_pt(H)
    page = out_doc.new_page(width=Wpt, height=Hpt)

    # 绘制纸张外框
    page.draw_rect(fitz.Rect(0, 0, Wpt, Hpt), color=(0, 0, 0), width=1.0)

    fontfile = pick_cjk_fontfile()
    ml = mm_to_pt(C.MARK_LEN)  # 刀线长度 5mm

    # 绘制每个图形
    for p in sheet["placements"]:
        tid = p["tid"]
        e = by_tid[tid]
        img_bytes = e["img_cont"] if is_contour else e["img_body"]

        x = float(p["x"])
        y = float(p["y"])
        w = float(p["w"])
        h = float(p["h"])

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

    # 整张纸标注：二维码固定在PDF右上角，标号固定在PDF左上角
    owner_tid = sheet.get("owner_tid", "")
    owner = by_tid.get(owner_tid, {})
    qr_text = sheet.get("qr_text") or owner.get("qr_text") or "MIX"
    qr_bytes = make_qr_png_bytes(qr_text)
    qr_rect = fitz.Rect(
        mm_to_pt(W - C.QR_W),
        mm_to_pt(0),
        mm_to_pt(W),
        mm_to_pt(C.QR_H),
    )
    page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
    if qr_bytes:
        page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

    lab = owner.get("label", "")
    if lab:
        xpt = mm_to_pt(1.0)
        ypt = mm_to_pt(C.QR_H - 1.5)
        max_w_pt = mm_to_pt(max(10.0, W - C.QR_W - 2.0))
        txt = trim_text_to_width(lab, C.LABEL_FONT_SIZE, max_w_pt)
        if fontfile:
            page.insert_text(fitz.Point(xpt, ypt), txt,
                             fontsize=C.LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
        else:
            page.insert_text(fitz.Point(xpt, ypt), txt,
                             fontsize=C.LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))

    # ---- 刀线：只在左侧边缘和顶部边缘，长度5mm ----
    col_edges = sheet.get("col_edges", [])
    row_edges = sheet.get("row_edges", [])

    # 顶部边缘的垂直刀线标记（列分割线）
    for xx in col_edges:
        xpt = mm_to_pt(xx)
        if 0 <= xpt <= Wpt:
            page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, ml),
                           color=(0, 0, 0), width=1.0)

    # 左侧边缘的水平刀线标记（行分割线）
    for yy in row_edges:
        ypt = mm_to_pt(yy)
        if 0 <= ypt <= Hpt:
            page.draw_line(fitz.Point(0, ypt), fitz.Point(ml, ypt),
                           color=(0, 0, 0), width=1.0)
