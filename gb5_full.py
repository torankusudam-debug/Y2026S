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

    sheets = []

    while any(remaining_global[e["tid"]] > 0 for e in entries):
        # 本张sheet的块列表和行信息
        sheet_blocks = []      # 每个块的完整信息（含位置）
        sheet_placements = []  # 每个图形的位置
        col_edges = set()
        row_edges = set()

        y_cursor = 0.0
        any_placed_this_sheet = False

        while y_cursor < C.FULL_H_MAX:
            # 开始新的一行
            row_blocks = []
            x_cursor = 0.0
            row_h = 0.0

            for e in entries:
                tid = e["tid"]
                if remaining_global.get(tid, 0) <= 0:
                    continue

                block = _make_block_for_entry(e, remaining_global[tid])
                if block is None:
                    continue

                bw = block["block_w"]
                bh = block["block_h"]

                # 检查是否能放入当前行
                needed_w = bw if x_cursor == 0 else (C.BLOCK_GAP + bw)
                if x_cursor + needed_w > C.FULL_W_MAX + 0.01:
                    continue

                # 检查行高是否会超过纸张高度
                new_row_h = max(row_h, bh)
                if y_cursor + new_row_h > C.FULL_H_MAX + 0.01:
                    continue

                actual_x = x_cursor + (C.BLOCK_GAP if x_cursor > 0 else 0)

                row_blocks.append({
                    "block": block,
                    "x": actual_x,
                    "tid": tid,
                })
                x_cursor = actual_x + bw
                row_h = new_row_h

                # 更新剩余数量
                remaining_global[tid] -= block["placed"]
                if remaining_global[tid] < 0:
                    remaining_global[tid] = 0

            if not row_blocks:
                break

            # 统一行高：同一行所有块使用相同行高（取最大值）
            actual_row_h = row_h

            row_edges.add(round(y_cursor, 2))

            for rb in row_blocks:
                block = rb["block"]
                bx = rb["x"]
                by = y_cursor
                bw = block["block_w"]
                bh = actual_row_h  # 使用统一行高

                col_edges.add(round(bx, 2))
                col_edges.add(round(bx + bw, 2))

                # 记录块信息（含位置）
                block_info = {
                    "tid": rb["tid"],
                    "x": bx,
                    "y": by,
                    "block_w": bw,
                    "block_h": block["block_h"],  # 实际块高
                    "display_h": bh,               # 显示行高（统一）
                    "rotated": block["rotated"],
                    "cols": block["cols"],
                    "rows": block["rows"],
                    "img_w": block["img_w"],
                    "img_h": block["img_h"],
                    "placed": block["placed"],
                }
                sheet_blocks.append(block_info)

                # 生成块内每个图形的位置
                placed_so_far = 0
                for row_i in range(block["rows"]):
                    for col_i in range(block["cols"]):
                        if placed_so_far >= block["placed"]:
                            break
                        px = bx + col_i * block["img_w"]
                        py = by + row_i * block["img_h"]
                        sheet_placements.append({
                            "tid": rb["tid"],
                            "x": px,
                            "y": py,
                            "w": block["img_w"],
                            "h": block["img_h"],
                            "rot": block["rotated"],
                            "block_idx": len(sheet_blocks) - 1,
                        })
                        placed_so_far += 1
                    if placed_so_far >= block["placed"]:
                        break

                any_placed_this_sheet = True

            row_edges.add(round(y_cursor + actual_row_h, 2))
            y_cursor += actual_row_h + C.BLOCK_GAP

        if not any_placed_this_sheet:
            # 无法放入任何图形
            problem = [e["tid"] for e in entries if remaining_global.get(e["tid"], 0) > 0]
            if problem:
                raise RuntimeError("全拼无法放入图形：%s" % str(problem))
            break

        # 计算实际使用的宽高
        W_used = C.FULL_W_MAX  # 固定600mm
        H_used = y_cursor - C.BLOCK_GAP if y_cursor > C.BLOCK_GAP else y_cursor

        # 确定owner
        placed_count_by_tid = {}
        for p in sheet_placements:
            tid = p["tid"]
            placed_count_by_tid[tid] = placed_count_by_tid.get(tid, 0) + 1
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
        _log(log_cb, "🧩 网格全拼出一张：owner=%s W=%d H=%d blocks=%d placed=%d" %
             (owner_tid, sheet["W"], sheet["H"], len(sheet_blocks), len(sheet_placements)))

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

    # 每个块标注二维码（左上角）和单号（右上角）
    for blk in sheet["blocks"]:
        tid = blk["tid"]
        e = by_tid[tid]
        bx = float(blk["x"])
        by_pos = float(blk["y"])
        bw = float(blk["block_w"])

        # ---- 二维码放块的左上角 ----
        qr_text = e.get("qr_text", tid)
        qr_bytes = make_qr_png_bytes(qr_text)
        qr_rect = fitz.Rect(
            mm_to_pt(bx),
            mm_to_pt(by_pos),
            mm_to_pt(bx + C.QR_W),
            mm_to_pt(by_pos + C.QR_H),
        )
        page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
        if qr_bytes:
            page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

        # ---- 单号放块的右上角 ----
        lab = e.get("label", "")
        if lab:
            max_label_w_mm = max(5.0, bw - C.QR_W - 2.0)
            # 单号在右上角
            xpt = mm_to_pt(bx + bw - max_label_w_mm)
            ypt = mm_to_pt(by_pos + C.QR_H - 1.5)
            max_w_pt = mm_to_pt(max_label_w_mm)
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
        if xpt > 0 and xpt < Wpt:
            page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, ml),
                           color=(0, 0, 0), width=1.0)

    # 左侧边缘的水平刀线标记（行分割线）
    for yy in row_edges:
        ypt = mm_to_pt(yy)
        if ypt > 0 and ypt < Hpt:
            page.draw_line(fitz.Point(0, ypt), fitz.Point(ml, ypt),
                           color=(0, 0, 0), width=1.0)
