# -*- coding: utf-8 -*-
"""
gb5_single.py —— 整拼模块

整拼标注规则：
  - 二维码放右上角
  - 单号放左上角
  - 刀线只标注在左侧边缘和顶部边缘
  - 320×464为图形的一块
"""
import math
import fitz

import gb5_config as C
from gb5_utils import mm_to_pt, pick_cjk_fontfile, trim_text_to_width


def cap_block(dim_mm, block_max_mm):
    dim = float(dim_mm)
    if dim <= 0:
        return 0
    cap = int(block_max_mm // dim)
    return cap if cap >= 1 else 0


def groups_for_k(capx, k):
    m = int(math.ceil(k / float(capx)))
    groups = [capx] * (m - 1)
    last = k - (m - 1) * capx
    if last <= 0 or last > capx:
        return None, None
    groups.append(last)
    return m, groups


def max_rows_fit(h_mm, capy, H_sheet):
    avail = float(H_sheet) - C.QR_BAND
    h = float(h_mm)
    if avail < h:
        return 0, 0.0

    best_rows = 0
    best_s = 0
    for s in range(1, 10000):
        avail_rows_height = avail - C.GAP * (s - 1)
        if avail_rows_height < h:
            break
        rows_fit = int(avail_rows_height // h)
        if rows_fit <= (s - 1) * capy:
            continue
        rows = min(rows_fit, s * capy)
        if rows > best_rows:
            best_rows = rows
            best_s = s

    usedH = best_rows * h + C.GAP * (best_s - 1)
    return int(best_rows), float(usedH)


def solve_single_type_no_waste(outer_w, outer_h, need_count):
    best = None
    for ori in (0, 1):
        z = outer_w if ori == 0 else outer_h
        h = outer_h if ori == 0 else outer_w

        capx = cap_block(z, C.BLOCK_W)
        block_max_h = C.BLOCK_H
        capy = cap_block(h, block_max_h)
        if capx <= 0 or capy <= 0:
            continue

        k_upper = int(C.SINGLE_W_MAX // z) + 3
        for k in range(1, k_upper + 1):
            m, groups = groups_for_k(capx, k)
            if groups is None:
                continue
            R = k * z + C.GAP * (m - 1)
            if R < C.SINGLE_W_MIN - 1e-9:
                continue
            if R > C.SINGLE_W_MAX + 1e-9:
                break

            for H_sheet in range(int(C.SINGLE_H_MIN), int(C.SINGLE_H_MAX) + 1):
                rows_max, used_parts_h = max_rows_fit(h, capy, H_sheet)
                if rows_max <= 0:
                    continue
                H_real = C.QR_BAND + used_parts_h
                if H_real > C.SINGLE_H_MAX + 1e-9:
                    continue

                cap_sheet = rows_max * k
                pages = int(math.ceil(need_count / float(cap_sheet)))
                area = pages * (R * H_real)

                cand = {
                    "ori": ori, "z": float(z), "h": float(h),
                    "k": int(k), "groups": list(groups),
                    "rows": int(rows_max),
                    "W": float(R),
                    "H": float(H_real),
                    "cap_sheet": int(cap_sheet),
                    "pages": int(pages),
                    "total_area": float(area),
                }

                if best is None or cand["total_area"] < best["total_area"] - 1e-9:
                    best = cand
                elif best is not None and abs(cand["total_area"] - best["total_area"]) <= 1e-9:
                    if cand["H"] < best["H"] - 1e-9:
                        best = cand

    return best


def build_single_placements(type_id, best):
    z = best["z"]
    h = best["h"]
    capy = int(math.floor(C.BLOCK_H / float(h)))
    if capy <= 0:
        capy = 1

    placements = []
    rows = best["rows"]
    groups = best["groups"]

    for r in range(rows):
        seg_idx = r // capy
        y = C.QR_BAND + r * h + seg_idx * C.GAP
        x = 0.0
        for gi, gsz in enumerate(groups):
            for _ in range(gsz):
                placements.append({"type": type_id, "x": x, "y": y, "w": z, "h": h, "rot": (best["ori"] == 1)})
                x += z
            if gi != len(groups) - 1:
                x += C.GAP
    return placements


def compute_edges_from_placements(placements):
    xs = set()
    ys = set()
    for p in placements:
        xs.add(int(round(p["x"])))
        xs.add(int(round(p["x"] + p["w"])))
        ys.add(int(round(p["y"])))
        ys.add(int(round(p["y"] + p["h"])))
    return sorted(xs), sorted(ys)


def draw_label_in_band(page, text, x_mm, y_mm, w_mm):
    """在左上角绘制单号标签。"""
    fontfile = pick_cjk_fontfile()
    xpt = mm_to_pt(x_mm + 1.0)
    ypt = mm_to_pt(y_mm - 1.5)
    max_w = mm_to_pt(w_mm - C.QR_W - 2.0)
    txt = trim_text_to_width(text, C.LABEL_FONT_SIZE, max_w)

    if fontfile:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=C.LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=C.LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))


def append_pages_single(out_doc, best, placements, pages, qr_bytes, label_text, img_bytes):
    """
    整拼绘制页面。

    标注规则：
      - 二维码放右上角
      - 单号放左上角
      - 刀线只标注在左侧边缘和顶部边缘
    """
    W = best["W"]
    H = best["H"]
    Wpt = mm_to_pt(W)
    Hpt = mm_to_pt(H)

    # 二维码放右上角
    qr_rect = fitz.Rect(Wpt - mm_to_pt(C.QR_W), 0, Wpt, mm_to_pt(C.QR_H))

    x_edges, y_edges = compute_edges_from_placements(placements)
    ml = mm_to_pt(C.MARK_LEN)

    # 找到左上角第一个placement（用于标注单号位置）
    first_p = None
    for p in placements:
        if first_p is None or p["y"] < first_p["y"] - 1e-9 or (abs(p["y"] - first_p["y"]) <= 1e-9 and p["x"] < first_p["x"]):
            first_p = p

    for _ in range(pages):
        page = out_doc.new_page(width=Wpt, height=Hpt)

        # 纸张外框
        page.draw_rect(fitz.Rect(0, 0, Wpt, Hpt), color=(0, 0, 0), width=1.0)

        # 二维码放右上角
        page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
        if qr_bytes:
            page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

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
                page.insert_image(inner, stream=img_bytes, keep_proportion=True, rotate=(90 if p["rot"] else 0))

        # 单号放左上角
        if first_p is not None and label_text:
            draw_label_in_band(page, label_text, first_p["x"], first_p["y"], first_p["w"])

        # ---- 刀线：只在顶部边缘（垂直刀线标记）和左侧边缘（水平刀线标记） ----
        # 顶部边缘：垂直刀线
        for xx in x_edges:
            xpt = mm_to_pt(xx)
            if xpt > 0 and xpt < Wpt:
                page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, ml), color=(0, 0, 0), width=1.0)

        # 左侧边缘：水平刀线
        for yy in y_edges:
            ypt = mm_to_pt(yy)
            if ypt > 0 and ypt < Hpt:
                page.draw_line(fitz.Point(0, ypt), fitz.Point(ml, ypt), color=(0, 0, 0), width=1.0)
