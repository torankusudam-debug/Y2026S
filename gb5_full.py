# -*- coding: utf-8 -*-
"""
gb5_full.py —— 全拼（混拼）模块

核心逻辑：
  - 网格排列：所有图形按行列网格放置
  - 通栏切割：先垂直切（从左到右），再水平切（从上到下）
  - 每一刀都是贯穿整张纸的，不会误伤图形
  - 同一列的所有图形宽度相同，同一行的所有图形高度相同
  - 标注：单号放图右上角，二维码放图左上角
  - 刀线只标注在左侧边缘和顶部边缘
"""

import fitz
import gb5_config as C
from gb5_utils import mm_to_pt, _log, pick_cjk_fontfile, trim_text_to_width
from gb5_qr import make_qr_png_bytes


def _try_grid_layout(entries_with_need, W_max, H_max, gap, label_band):
    """
    尝试将 entries 按网格排列到一张纸上。

    网格规则：
      - 每列宽度 = 该列中最宽图形的宽度
      - 每行高度 = 该行中最高图形的高度 + label_band（标注区域）
      - 列间距 = gap，行间距 = gap
      - 先从左到右填列，再从上到下填行
      - 保证通栏切割：同一列宽度一致，同一行高度一致

    由于不同图形尺寸可能不同，我们采用"贪心行填充"策略：
      - 对于每一行，从左到右放置图形，累加宽度不超过 W_max
      - 同一行中所有图形的行高取最大值
      - 行高累加不超过 H_max

    返回:
      placements: list of dict
      col_edges: 列的x坐标边界（用于垂直刀线）
      row_edges: 行的y坐标边界（用于水平刀线）
      W_used, H_used: 实际使用的宽高
      placed_count: dict {tid: 已放置数量}
    """
    placements = []
    col_edges_set = set()
    row_edges_set = set()
    placed_count = {}  # tid -> count

    # 构建待放置列表：(entry, remaining)
    remaining = {}
    for e in entries_with_need:
        tid = e["tid"]
        remaining[tid] = int(e["need"])

    y_cursor = 0.0
    all_done = False

    while not all_done and y_cursor < H_max:
        # 开始新的一行
        row_items = []  # [(entry, x, w_cell, h_cell, rotated)]
        x_cursor = 0.0
        row_h = 0.0

        for e in entries_with_need:
            tid = e["tid"]
            if remaining.get(tid, 0) <= 0:
                continue

            ow = float(e["ow"])
            oh = float(e["oh"])

            # 尝试不旋转和旋转两种方式
            # 单元格宽 = 图形宽 + gap（最后一列不加gap，但为简化先加上）
            # 单元格高 = 图形高 + label_band
            candidates = []
            # 不旋转
            cell_w = ow
            cell_h = oh + label_band
            candidates.append((cell_w, cell_h, False, ow, oh))
            # 旋转90度
            if int(ow) != int(oh):
                cell_w_r = oh
                cell_h_r = ow + label_band
                candidates.append((cell_w_r, cell_h_r, True, oh, ow))

            placed_any = True
            while placed_any and remaining.get(tid, 0) > 0:
                placed_any = False
                for cell_w, cell_h, rotated, w_img, h_img in candidates:
                    needed_w = cell_w if x_cursor == 0 else (gap + cell_w)
                    if x_cursor + needed_w > W_max + 0.01:
                        continue
                    needed_h = cell_h if y_cursor == 0 else cell_h
                    if y_cursor + max(row_h, needed_h) > H_max + 0.01:
                        continue

                    actual_x = x_cursor + (gap if x_cursor > 0 else 0)
                    row_items.append({
                        "entry": e,
                        "x": actual_x,
                        "w_img": w_img,
                        "h_img": h_img,
                        "cell_w": cell_w,
                        "cell_h": cell_h,
                        "rotated": rotated,
                    })
                    x_cursor = actual_x + cell_w
                    row_h = max(row_h, cell_h)
                    remaining[tid] -= 1
                    placed_any = True
                    break

        if not row_items:
            break

        # 统一行高：同一行所有图形使用相同行高
        actual_row_h = row_h
        if y_cursor + actual_row_h > H_max + 0.01:
            break

        row_top = y_cursor
        row_edges_set.add(round(row_top, 2))

        for ri in row_items:
            e = ri["entry"]
            tid = e["tid"]
            x = ri["x"]
            w_img = ri["w_img"]
            h_img = ri["h_img"]

            # 图形放在单元格内，标注区在上方
            img_y = row_top + label_band
            placements.append({
                "tid": tid,
                "x": x,
                "y": img_y,
                "w": w_img,
                "h": h_img,
                "rot": ri["rotated"],
                "label_y": row_top,  # 标注区域顶部y
                "cell_w": ri["cell_w"],
            })

            col_edges_set.add(round(x, 2))
            col_edges_set.add(round(x + ri["cell_w"], 2))

            if tid not in placed_count:
                placed_count[tid] = 0
            placed_count[tid] += 1

        row_edges_set.add(round(row_top + actual_row_h, 2))
        y_cursor = row_top + actual_row_h + gap

        # 检查是否全部放完
        all_done = all(remaining.get(e["tid"], 0) <= 0 for e in entries_with_need)

    W_used = max(C.FULL_W_MIN, min(C.FULL_W_MAX, max((p["x"] + p["cell_w"]) for p in placements) if placements else 0))
    H_used = min(H_max, y_cursor - gap if y_cursor > gap else y_cursor) if placements else 0

    col_edges = sorted(col_edges_set)
    row_edges = sorted(row_edges_set)

    return placements, col_edges, row_edges, float(W_used), float(H_used), placed_count, remaining


def pack_full_grid_sheets(entries, log_cb=None):
    """
    将全拼条目按网格排列打包成多张sheet。

    每张sheet保证：
      - 宽度不超过 FULL_W_MAX (600mm)
      - 高度不超过 FULL_H_MAX (1500mm)
      - 网格排列，支持通栏切割

    返回:
      sheets: list of dict
      by_tid: dict {tid: entry}
    """
    entries = list(entries)
    by_tid = {}
    for e in entries:
        by_tid[e["tid"]] = e

    # 构建待处理列表（带剩余数量）
    remaining_global = {}
    for e in entries:
        remaining_global[e["tid"]] = int(e["need"])

    sheets = []

    while any(remaining_global[e["tid"]] > 0 for e in entries):
        # 构建本轮待放置条目
        entries_this_round = []
        for e in entries:
            tid = e["tid"]
            if remaining_global[tid] > 0:
                entries_this_round.append({
                    "tid": tid,
                    "ow": e["ow"],
                    "oh": e["oh"],
                    "need": remaining_global[tid],
                })

        if not entries_this_round:
            break

        placements, col_edges, row_edges, W_used, H_used, placed_count, remaining_after = \
            _try_grid_layout(
                entries_this_round,
                C.FULL_W_MAX,
                C.FULL_H_MAX - C.FULL_TOP_PAD,
                C.GAP,
                C.FULL_LABEL_BAND,
            )

        if not placements:
            # 无法放入任何图形，报错
            problem_entries = [e["tid"] for e in entries_this_round if remaining_global[e["tid"]] > 0]
            raise RuntimeError("全拼无法放入图形：%s" % str(problem_entries))

        # 更新全局剩余
        for tid, cnt in placed_count.items():
            remaining_global[tid] -= cnt
            if remaining_global[tid] < 0:
                remaining_global[tid] = 0

        # 确定owner（放置数量最多的类型）
        owner_tid = max(placed_count, key=placed_count.get)
        owner_qr = by_tid[owner_tid].get("qr_text", "MIX")

        sheet = {
            "owner_tid": owner_tid,
            "qr_text": owner_qr,
            "W": int(max(C.FULL_W_MIN, W_used)),
            "H": int(min(C.FULL_H_MAX, C.FULL_TOP_PAD + H_used)),
            "placements": placements,
            "col_edges": col_edges,
            "row_edges": row_edges,
        }
        sheets.append(sheet)
        _log(log_cb, "🧩 网格全拼出一张：owner=%s W=%d H=%d placed=%d" %
             (owner_tid, sheet["W"], sheet["H"], len(placements)))

    return sheets, by_tid


def append_one_full_sheet(out_doc, sheet, by_tid, is_contour=False):
    """
    将一个全拼sheet绘制到out_doc中（新增一页）。

    标注规则：
      - 单号放图右上角
      - 二维码放图左上角
      - 刀线只标注在左侧边缘和顶部边缘
    """
    W = int(sheet["W"])
    H = int(sheet["H"])
    Wpt = mm_to_pt(W)
    Hpt = mm_to_pt(H)
    page = out_doc.new_page(width=Wpt, height=Hpt)

    # 绘制纸张外框
    page.draw_rect(fitz.Rect(0, 0, Wpt, Hpt), color=(0, 0, 0), width=1.0)

    fontfile = pick_cjk_fontfile()
    ml = mm_to_pt(C.MARK_LEN)

    # 绘制每个图形
    for p in sheet["placements"]:
        tid = p["tid"]
        e = by_tid[tid]
        img_bytes = e["img_cont"] if is_contour else e["img_body"]

        x = float(p["x"])
        y = float(p["y"]) + C.FULL_TOP_PAD
        w = float(p["w"])
        h = float(p["h"])
        label_y = float(p["label_y"]) + C.FULL_TOP_PAD
        cell_w = float(p["cell_w"])

        # 绘制图形
        outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
        if C.DRAW_PART_OUTER_BOX:
            page.draw_rect(outer, color=(0, 0, 0), width=0.5)

        inner = fitz.Rect(
            outer.x0 + mm_to_pt(C.INNER_MARGIN_MM),
            outer.y0 + mm_to_pt(C.INNER_MARGIN_MM),
            outer.x1 - mm_to_pt(C.INNER_MARGIN_MM),
            outer.y1 - mm_to_pt(C.INNER_MARGIN_MM),
        )
        page.insert_image(inner, stream=img_bytes, keep_proportion=True,
                          rotate=(90 if p["rot"] else 0))

        # ---- 标注：二维码放图左上角 ----
        qr_text = e.get("qr_text", tid)
        qr_bytes = make_qr_png_bytes(qr_text)
        qr_rect = fitz.Rect(
            mm_to_pt(x),
            mm_to_pt(label_y),
            mm_to_pt(x + C.QR_W),
            mm_to_pt(label_y + C.QR_H),
        )
        page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
        page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

        # ---- 标注：单号放图右上角 ----
        lab = e.get("label", "")
        if lab:
            max_label_w_mm = max(5.0, cell_w - C.QR_W - 2.0)
            # 单号在右上角：x起点在二维码右边
            xpt = mm_to_pt(x + C.QR_W + 1.0)
            ypt = mm_to_pt(label_y + C.QR_H - 1.5)
            max_w_pt = mm_to_pt(max_label_w_mm)
            txt = trim_text_to_width(lab, C.LABEL_FONT_SIZE, max_w_pt)
            if fontfile:
                page.insert_text(fitz.Point(xpt, ypt), txt,
                                 fontsize=C.LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
            else:
                page.insert_text(fitz.Point(xpt, ypt), txt,
                                 fontsize=C.LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))

    # ---- 刀线：只在左侧边缘和顶部边缘 ----
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
        ypt = mm_to_pt(yy + C.FULL_TOP_PAD)
        if ypt > 0 and ypt < Hpt:
            page.draw_line(fitz.Point(0, ypt), fitz.Point(ml, ypt),
                           color=(0, 0, 0), width=1.0)
