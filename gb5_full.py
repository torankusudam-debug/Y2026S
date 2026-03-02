# -*- coding: utf-8 -*-
import fitz
import gb5_config as C
from gb5_utils import mm_to_pt, _log, pick_cjk_fontfile, trim_text_to_width
from gb5_qr import make_qr_png_bytes


class MaxRectsBin(object):
    def __init__(self, width, height):
        self.width = int(width)
        self.height = int(height)
        self.free_rects = [(0, 0, self.width, self.height)]
        self.used_width = 0

    def _score(self, fw, fh, rw, rh, x, y, mode):
        area_fit = (fw * fh) - (rw * rh)
        short_side = min(abs(fw - rw), abs(fh - rh))
        long_side = max(abs(fw - rw), abs(fh - rh))
        if mode == 0:
            return (area_fit, short_side, long_side, y, x)
        if mode == 1:
            return (short_side, long_side, area_fit, y, x)
        return (y, x, area_fit, short_side, long_side)

    def find_position(self, rw, rh, mode=0):
        best = None
        for fx, fy, fw, fh in self.free_rects:
            if rw <= fw and rh <= fh:
                score = self._score(fw, fh, rw, rh, fx, fy, mode)
                if (best is None) or (score < best["score"]):
                    best = {"x": fx, "y": fy, "w": rw, "h": rh, "score": score}
        return best

    def place(self, p):
        x = int(p["x"]); y = int(p["y"]); w = int(p["w"]); h = int(p["h"])
        new_free = []
        for fx, fy, fw, fh in self.free_rects:
            if (x >= fx + fw) or (x + w <= fx) or (y >= fy + fh) or (y + h <= fy):
                new_free.append((fx, fy, fw, fh))
                continue

            if y > fy:
                new_free.append((fx, fy, fw, y - fy))
            if y + h < fy + fh:
                new_free.append((fx, y + h, fw, (fy + fh) - (y + h)))

            top = max(fy, y)
            bot = min(fy + fh, y + h)
            hh = bot - top
            if hh > 0 and x > fx:
                new_free.append((fx, top, x - fx, hh))
            if hh > 0 and x + w < fx + fw:
                new_free.append((x + w, top, (fx + fw) - (x + w), hh))

        self.free_rects = self._prune(new_free)
        xw = x + w
        if xw > self.used_width:
            self.used_width = xw

    def _prune(self, rects):
        cleaned = []
        for x, y, w, h in rects:
            if w <= 0 or h <= 0:
                continue
            if x < 0 or y < 0:
                continue
            if x + w > self.width or y + h > self.height:
                continue
            cleaned.append((int(x), int(y), int(w), int(h)))

        if not cleaned:
            return []

        cleaned.sort(key=lambda t: (-(t[2] * t[3]), t[1], t[0]))

        out = []
        for xi, yi, wi, hi in cleaned:
            contained = False
            for xj, yj, wj, hj in out:
                if xi >= xj and yi >= yj and xi + wi <= xj + wj and yi + hi <= yj + hj:
                    contained = True
                    break
            if not contained:
                out.append((xi, yi, wi, hi))

        out.sort(key=lambda t: (t[1], t[0], -(t[2] * t[3])))
        return out


def _best_pos_for_item(bp, ow, oh):
    modes = (2, 0, 1)
    best = None
    candidates = [(int(ow + C.GAP), int(oh + C.FULL_LABEL_BAND), False, int(ow), int(oh))]
    if int(ow) != int(oh):
        candidates.append((int(oh + C.GAP), int(ow + C.FULL_LABEL_BAND), True, int(oh), int(ow)))

    for mode in modes:
        for pw, ph, rot, w_img, h_img in candidates:
            pos = bp.find_position(pw, ph, mode=mode)
            if pos is None:
                continue
            score = pos.get("score", (0, 0, 0, 0, 0))
            if (best is None) or (score < best["score"]):
                best = {"pos": pos, "rot": rot, "pw": pw, "ph": ph, "w": w_img, "h": h_img, "score": score}
    return best


def pack_full_sequential_sheets(entries, log_cb=None):
    entries = list(entries)
    remaining = {}
    by_tid = {}
    for e in entries:
        tid = e["tid"]
        by_tid[tid] = e
        remaining[tid] = int(e["need"])

    sheets = []
    cur_idx = 0

    while cur_idx < len(entries):
        while cur_idx < len(entries) and remaining[entries[cur_idx]["tid"]] <= 0:
            cur_idx += 1
        if cur_idx >= len(entries):
            break

        owner_tid = entries[cur_idx]["tid"]
        owner_qr = by_tid[owner_tid].get("qr_text", "MIX")

        binW = C.FULL_W_MAX
        binH = C.FULL_H_MAX - int(C.FULL_TOP_PAD)
        if binH <= 0:
            binH = 1

        bp = MaxRectsBin(binW, binH)
        placements = []
        first_pos = {}
        used_w = 0
        used_h = 0
        type_ptr = cur_idx

        while True:
            while type_ptr < len(entries) and remaining[entries[type_ptr]["tid"]] <= 0:
                type_ptr += 1
            if type_ptr >= len(entries):
                break

            e = entries[type_ptr]
            tid = e["tid"]

            best = _best_pos_for_item(bp, e["ow"], e["oh"])
            if best is None:
                if tid == owner_tid and remaining[tid] > 0:
                    break
                type_ptr += 1
                continue

            pos = best["pos"]
            bp.place(pos)

            x_box = int(pos["x"])
            y_box = int(pos["y"])
            w_img = int(best["w"])
            h_img = int(best["h"])

            placements.append({"tid": tid, "x": x_box, "y": y_box, "w": w_img, "h": h_img, "rot": bool(best["rot"])})
            if tid not in first_pos:
                first_pos[tid] = {"x": x_box, "y": y_box, "w": w_img, "h": h_img, "rot": bool(best["rot"])}

            remaining[tid] -= 1
            used_w = max(used_w, x_box + w_img)
            used_h = max(used_h, y_box + int(C.FULL_LABEL_BAND) + h_img)

            if tid == owner_tid and remaining[tid] <= 0:
                while cur_idx < len(entries) and remaining[entries[cur_idx]["tid"]] <= 0:
                    cur_idx += 1
                type_ptr = cur_idx

        if not placements:
            raise RuntimeError("全拼也无法放入：%s (ow=%s,oh=%s)" % (owner_tid, str(by_tid[owner_tid]["ow"]), str(by_tid[owner_tid]["oh"])))

        W_sheet = int(max(C.FULL_W_MIN, min(C.FULL_W_MAX, used_w)))
        H_sheet = int(min(C.FULL_H_MAX, int(C.FULL_TOP_PAD) + used_h))
        if H_sheet < 1:
            H_sheet = 1

        sheets.append({"owner_tid": owner_tid, "qr_text": owner_qr, "W": W_sheet, "H": H_sheet, "placements": placements, "first_pos": first_pos})
        _log(log_cb, "🧩 顺序全拼出一张：owner=%s  W=%d H=%d  placed=%d" % (owner_tid, W_sheet, H_sheet, len(placements)))

    return sheets, by_tid


def append_one_full_sheet(out_doc, sheet, by_tid, is_contour=False):
    W = int(sheet["W"]); H = int(sheet["H"])
    Wpt = mm_to_pt(W); Hpt = mm_to_pt(H)
    page = out_doc.new_page(width=Wpt, height=Hpt)

    page.draw_rect(fitz.Rect(0, 0, Wpt, Hpt), color=(0, 0, 0), width=1.0)

    for p in sheet["placements"]:
        tid = p["tid"]
        e = by_tid[tid]
        img_bytes = e["img_cont"] if is_contour else e["img_body"]

        x = float(p["x"])
        y = float(C.FULL_TOP_PAD + p["y"] + C.FULL_LABEL_BAND)
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
        page.insert_image(inner, stream=img_bytes, keep_proportion=True, rotate=(90 if p["rot"] else 0))

    fontfile = pick_cjk_fontfile()
    for tid, fp in sheet["first_pos"].items():
        e = by_tid[tid]
        lab = e.get("label", "")
        if not lab:
            continue

        qr_text = e.get("qr_text", tid)
        qr_bytes = make_qr_png_bytes(qr_text)

        x = float(fp["x"])
        w = float(fp["w"])

        y_band_top = float(C.FULL_TOP_PAD + fp["y"] + C.FULL_LABEL_BAND)
        y_band_bottom = y_band_top - float(C.FULL_LABEL_BAND)

        qr_x0 = x + w - C.QR_W
        qr_rect = fitz.Rect(mm_to_pt(qr_x0), mm_to_pt(y_band_bottom), mm_to_pt(qr_x0 + C.QR_W), mm_to_pt(y_band_bottom + C.QR_H))
        page.draw_rect(qr_rect, color=(0, 0, 0), width=1.0)
        page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

        max_label_w_mm = max(5.0, w - C.QR_W - 2.0)
        xpt = mm_to_pt(x + 1.0)
        ypt = mm_to_pt(y_band_top - 1.5)
        max_w_pt = mm_to_pt(max_label_w_mm)

        txt = trim_text_to_width(lab, C.LABEL_FONT_SIZE, max_w_pt)
        if fontfile:
            page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=C.LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
        else:
            page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=C.LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))