# -*- coding: utf-8 -*-
"""
拼图/拼版脚本（双输出版：图形页 + 轮廓页）

✅ 本次更新（按你要求）：
1) 刀线位置规则改为：
   - 顶部刀线（短刀线刻度）：
        x1 = 5
        x_i = 5 + n*i
        若跨过 320（x_(i-1)<320 且 x_i>320），则从该 i 起整体 +10（等价于跳过 320~330 空白带）
   - 左侧刀线（短刀线刻度）：
        y1 = 10
        y_i = 10 + m*i
        若跨过 464，则从该 i 起整体 +10（等价于插入 10mm 空隙）
        若跨过 928(=464*2)，则从该 i 起整体 +20（再插入一段 10mm 空隙）
   其中 n/m 为图形宽高（旋转后以实际排布宽高为准）。

2) 删除所有全拼/混拼代码（*_mix.pdf 不再生成）。
3) 数量 N < 10：不拼版，直接跳过（不输出）。
"""

import os
import re
import math
import time
import shutil
import hashlib
from io import BytesIO

import fitz
import numpy as np
from PIL import Image

try:
    import cv2
    CV2_OK = True
except Exception:
    CV2_OK = False

try:
    import qrcode
    QR_OK = True
except Exception:
    QR_OK = False

try:
    from PIL import ImageDraw
    PIL_DRAW_OK = True
except Exception:
    PIL_DRAW_OK = False


# =========================
# 路径（可由 run.py 动态注入）
# =========================
DEST_DIR = r"D:\test_data\dest"
IN_PDF_ARCHIVE_DIR = r"D:\test_data\test"
DEST_DIR1 = r"D:\test_data\gest"
DEST_DIR2 = r"D:\test_data\pest"


def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)


# =========================
# 参数
# =========================
OUTER_EXT = 2.0
INNER_MARGIN_MM = 5.0

MARK_LEN = 5.0  # 刀线/刻度长度（mm）

SINGLE_W = 600.0
SINGLE_H_MAX = 1500.0

# 边距 & 320空白带
SINGLE_MARGIN_L = 5.0
SINGLE_MARGIN_R = 5.0
SINGLE_SPLIT_X = 320.0
SINGLE_SPLIT_GAP_W = 10.0

SINGLE_RESERVED_W = SINGLE_SPLIT_GAP_W + (SINGLE_MARGIN_L + SINGLE_MARGIN_R)  # 20
SINGLE_USABLE_W = SINGLE_W - SINGLE_RESERVED_W  # 580

# 单拼段右上 QR 区
QR_BAND = 10.0
QR_W = 10.0
QR_H = 10.0

LABEL_FONT_SIZE = 10
RENDER_DPI = 600
CROP_PAD_MM = 1.2
TEXT_MASK_PAD_MM = 1.2

DRAW_PART_OUTER_BOX = False
# =========================
# 基础工具
# =========================
def ensure_dir(p):
    if p and (not os.path.isdir(p)):
        os.makedirs(p)


def mm_to_pt(mm):
    return mm * 72.0 / 25.4


def mm_to_px(mm, dpi):
    return int(round(mm * dpi / 25.4))


def _log(log_cb, s):
    if log_cb:
        log_cb(s)
    else:
        print(s)


def clamp_bbox(x0, y0, x1, y1, img_w, img_h):
    x0 = max(0, min(img_w - 1, int(round(x0))))
    y0 = max(0, min(img_h - 1, int(round(y0))))
    x1 = max(x0 + 1, min(img_w, int(round(x1))))
    y1 = max(y0 + 1, min(img_h, int(round(y1))))
    return x0, y0, x1, y1


def union_bbox(b1, b2):
    if b1 is None:
        return b2
    if b2 is None:
        return b1
    return (
        min(b1[0], b2[0]),
        min(b1[1], b2[1]),
        max(b1[2], b2[2]),
        max(b1[3], b2[3]),
    )


def rect_union(rects):
    rects = [r for r in rects if r is not None]
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    return fitz.Rect(x0, y0, x1, y1)


def archive_input_pdf_to_dir(src_pdf_path, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src_pdf_path)
    dst_path = os.path.join(dst_dir, base)

    try:
        if os.path.abspath(src_pdf_path).lower() == os.path.abspath(dst_path).lower():
            return dst_path
    except Exception:
        pass

    if os.path.exists(dst_path):
        try:
            if os.path.getsize(dst_path) == os.path.getsize(src_pdf_path):
                return dst_path
        except Exception:
            pass
        name, ext = os.path.splitext(base)
        ts = time.strftime("%Y%m%d_%H%M%S")
        dst_path = os.path.join(dst_dir, "%s_%s%s" % (name, ts, ext))

    shutil.copy2(src_pdf_path, dst_path)
    return dst_path


def _sanitize_filename_for_windows(name_no_ext, max_len=160):
    s = str(name_no_ext)
    s = s.replace("\\", "_").replace("/", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace('"', "_")
    s = s.replace("<", "_").replace(">", "_").replace("|", "_")
    if len(s) <= max_len:
        return s
    h = hashlib.md5(s.encode("utf-8", "ignore")).hexdigest()[:10]
    keep = max(40, max_len - 11)
    return s[:keep] + "_" + h


# =========================
# 刀线绘制（5mm短刀线/刻度）
# =========================
def _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0):
    ypt = mm_to_pt(float(y_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(0, ypt), fitz.Point(tick, ypt), color=(0, 0, 0), width=lw)


def _draw_top_edge_tick(page, x_mm, tick_len_mm=MARK_LEN, lw=1.0):
    xpt = mm_to_pt(float(x_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, tick), color=(0, 0, 0), width=lw)

def _draw_top_tick_at_y(page, x_mm, y_mm, tick_len_mm=MARK_LEN, lw=1.0):
    """在 y=y_mm 这条水平线上，画一根向下的短刻度线"""
    xpt = mm_to_pt(float(x_mm))
    ypt = mm_to_pt(float(y_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, ypt), fitz.Point(xpt, ypt + tick), color=(0, 0, 0), width=lw)

# 建议放在参数区
EDGE_SNAP_MM = 10.0   # <=6mm 就吸到纸边（你现在通常是 5mm）


def _snap_edge_mm(v, lo, hi, eps):
    if abs(v - lo) <= eps:
        return lo
    if abs(v - hi) <= eps:
        return hi
    return v


def _draw_segment_cut_ticks(page, seg, yoff_mm, page_h_mm=SINGLE_H_MAX):
    """
    ✅ 每段单独画刀线刻度，并把“最外侧边界”吸到纸边：
      - 左外侧 x_min 若在 0~EDGE_SNAP_MM 内 -> x=0
      - 右外侧 x_max 若在 (W-EDGE_SNAP_MM)~W 内 -> x=W
      - 底外侧 y_max 若在 (H-EDGE_SNAP_MM)~H 内 -> y=H
    这样就能省掉外边那一刀。
    """
    pls = seg.get("placements") or []
    if not pls:
        return

    # 用 placements 推导网格（避免用错 n/m 导致乱标）
    w = float(pls[0].get("w", 0.0))
    h = float(pls[0].get("h", 0.0))
    if w <= 0 or h <= 0:
        return

    # 去重（浮点误差）
    xs = sorted({round(float(p["x"]), 3) for p in pls})
    ys = sorted({round(float(p["y"]), 3) for p in pls})

    # 生成所有竖边界：列左边界 + 列右边界
    x_lines = set(xs)
    for x in xs:
        x_lines.add(round(x + w, 3))
    x_lines = sorted(x_lines)

    if not x_lines:
        return

    W = float(SINGLE_W)
    x_min = float(min(x_lines))
    x_max = float(max(x_lines))

    # 段顶部(y=yoff)：画 x 刻度（向下 5mm）
    for x in x_lines:
        x_draw = float(x)
        # ✅ 外侧吸边（省一刀）
        if abs(x - x_min) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        elif abs(x - x_max) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)

        if 0.0 - 1e-6 <= x_draw <= W + 1e-6:
            _draw_top_tick_at_y(page, x_draw, yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)

    # 生成所有横边界：行上边界 + 行下边界
    y_lines = set(ys)
    for y in ys:
        y_lines.add(round(y + h, 3))
    y_lines = sorted(y_lines)

    if not y_lines:
        return

    y_min = float(min(y_lines))
    y_max = float(max(y_lines))

    # 段左侧(x=0)：画 y 刻度（向右 5mm）
    for y in y_lines:
        Y = float(yoff_mm) + float(y)
        # ✅ 只对“段的最底外侧边界”吸到整页底边（省一刀）
        if abs(y - y_max) <= 1e-6:
            Y = _snap_edge_mm(Y, 0.0, float(page_h_mm), EDGE_SNAP_MM)

        if 0.0 - 1e-6 <= Y <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, Y, tick_len_mm=MARK_LEN, lw=1.0)


def _draw_horizontal_cutline(page, y_mm, lw=1.0):
    ypt = mm_to_pt(float(y_mm))
    Wpt = mm_to_pt(float(SINGLE_W))
    page.draw_line(fitz.Point(0, ypt), fitz.Point(Wpt, ypt), color=(0, 0, 0), width=lw)


def _draw_vertical_cutline(page, x_mm, y0_mm, y1_mm, lw=1.0):
    xpt = mm_to_pt(float(x_mm))
    y0 = mm_to_pt(float(y0_mm))
    y1 = mm_to_pt(float(y1_mm))
    page.draw_line(fitz.Point(xpt, y0), fitz.Point(xpt, y1), color=(0, 0, 0), width=lw)


def _draw_edge_ticks_by_nm(page, n_mm, m_mm, W_mm=SINGLE_W, H_mm=SINGLE_H_MAX):
    """
    ✅ 区域块规则（你这条消息的要求）
    - 宽度方向：每个区域块宽度不超过 320mm；块与块之间间距固定 10mm
      x1=5，默认每次 +n
      若出现 x < boundary 且 x+n > boundary（boundary 初始 320）：
         下一根改为 x+10（插入块间距）
         boundary += 320（进入下一块，防止重复插入）
      然后继续 +n

    - 高度方向：每个区域块高度不超过 464mm；块与块之间间距固定 10mm
      y1=10，默认每次 +m
      若跨过 boundary(初始 464)：下一根 = y+10；boundary += 464
    """
    try:
        n = float(n_mm)
        m = float(m_mm)
    except Exception:
        return
    if n <= 0 or m <= 0:
        return

    # -------- 顶部 x 刀线 --------
    x = 5.0
    boundary_x = 320.0
    guard = 0
    while x <= W_mm + 1e-6:
        _draw_top_edge_tick(page, x, tick_len_mm=MARK_LEN, lw=1.0)

        nxt = x + n

        # ✅ 只在当前块边界处插一次 10mm
        if x < boundary_x and nxt > boundary_x:
            nxt = x + 10.0
            boundary_x += 320.0  # 进入下一块（600宽一般只会触发一次）

        if nxt <= x + 1e-9:
            break
        x = nxt

        guard += 1
        if guard > 20000:
            break

    # -------- 左侧 y 刀线 --------
    y = 10.0
    boundary_y = 464.0
    guard = 0
    while y <= H_mm + 1e-6:
        _draw_left_edge_tick(page, y, tick_len_mm=MARK_LEN, lw=1.0)

        nxt = y + m

        # ✅ 每跨过一个 464 块，就插一次 10mm
        if y < boundary_y and nxt > boundary_y:
            nxt = y + 10.0
            boundary_y += 464.0

        if nxt <= y + 1e-9:
            break
        y = nxt

        guard += 1
        if guard > 20000:
            break
# =========================
# 文件名解析 A*B^N
# =========================
def parse_A_B_N_from_filename(pdf_path):
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    parts = base.split("^")

    size_part = None
    n_part = None

    if len(parts) >= 8:
        size_part = parts[6]
        n_part = parts[7]
    else:
        for i in range(len(parts) - 1):
            if re.search(r"\d+(\.\d+)?\s*[\*xX×✕]\s*\d+(\.\d+)?", parts[i]) and re.search(r"\d+", parts[i + 1]):
                size_part = parts[i]
                n_part = parts[i + 1]
                break
        if size_part is None or n_part is None:
            raise ValueError("无法从文件名解析 A×B 与 N：%s" % base)

    m = re.search(r"(\d+(\.\d+)?)\s*[\*xX×✕]\s*(\d+(\.\d+)?)", str(size_part))
    if not m:
        raise ValueError("尺寸段不是 A×B 格式：%s" % str(size_part))

    A = float(m.group(1))
    B = float(m.group(3))

    m2 = re.search(r"(\d+)", str(n_part))
    if not m2:
        raise ValueError("数量段不是数字：%s" % str(n_part))
    N = int(m2.group(1))

    if A <= 0 or B <= 0 or N <= 0:
        raise ValueError("解析到的 A/B/N 非法：A=%s B=%s N=%s (%s)" % (A, B, N, base))

    return A, B, N


def extract_label_text(pdf_path):
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    ps = base.split("^")
    if len(ps) >= 2:
        return "^".join(ps[:2])
    return base[:40]


def extract_qr_text_from_filename(pdf_path):
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    ps = base.split("^")
    if len(ps) >= 10 and ps[9]:
        return str(ps[9]).strip()

    m = re.search(r"(SJ[0-9A-Za-z]+)", base)
    if m:
        return m.group(1)

    if len(ps) >= 2:
        return "^".join(ps[:2])

    return base[:80]


def make_qr_png_bytes(qr_text, box_px=240):
    if QR_OK:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=0
        )
        qr.add_data(qr_text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img = img.resize((box_px, box_px))
        bio = BytesIO()
        img.save(bio, format="PNG", optimize=True)
        return bio.getvalue()

    img = Image.new("RGB", (box_px, box_px), (255, 255, 255))
    if PIL_DRAW_OK:
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, box_px - 1, box_px - 1], outline=(0, 0, 0), width=3)
        d.text((box_px * 0.28, box_px * 0.40), "QR", fill=(0, 0, 0))
    bio = BytesIO()
    img.save(bio, format="PNG", optimize=True)
    return bio.getvalue()


def pick_cjk_fontfile():
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def trim_text_to_width(text, fontsize, max_width_pt):
    est = 0.55 * fontsize * len(text)
    if est <= max_width_pt:
        return text
    keep = max(0, int(max_width_pt / (0.55 * fontsize)) - 3)
    if keep <= 0:
        return "..."
    return text[:keep] + "..."


# =========================
# PDF结构识别 / bbox
# =========================
def _is_probably_page_frame(rect_obj, page_rect):
    if rect_obj is None:
        return False
    tol = 2.0
    if rect_obj.width <= 0 or rect_obj.height <= 0:
        return False
    full_like = (rect_obj.width >= 0.985 * page_rect.width and rect_obj.height >= 0.985 * page_rect.height)
    touches_all_edges = (
        abs(rect_obj.x0 - page_rect.x0) <= tol and
        abs(rect_obj.y0 - page_rect.y0) <= tol and
        abs(rect_obj.x1 - page_rect.x1) <= tol and
        abs(rect_obj.y1 - page_rect.y1) <= tol
    )
    return full_like and touches_all_edges


def get_page_struct_info(doc, page_index):
    page = doc.load_page(page_index)
    page_rect = page.rect

    text_boxes = []
    image_boxes = []

    try:
        td = page.get_text("dict")
        for b in td.get("blocks", []):
            btype = b.get("type", None)
            bb = b.get("bbox", None)
            if not bb:
                continue
            r = fitz.Rect(bb)
            if r.width <= 0 or r.height <= 0:
                continue
            if btype == 0:
                text_boxes.append(r)
            elif btype == 1:
                if not _is_probably_page_frame(r, page_rect):
                    image_boxes.append(r)
    except Exception:
        pass

    draw_rects = []
    try:
        drawings = page.get_drawings()
        for g in drawings:
            rr = g.get("rect", None)
            if rr is None:
                continue
            r = fitz.Rect(rr)
            if r.width <= 0.5 or r.height <= 0.5:
                continue
            if _is_probably_page_frame(r, page_rect):
                continue
            draw_rects.append(r)
    except Exception:
        pass

    return {
        "page_rect": page_rect,
        "text_boxes": text_boxes,
        "image_bbox": rect_union(image_boxes),
        "draw_bbox": rect_union(draw_rects),
    }


def render_page_to_pil(pdf_path, page_index=0, dpi=RENDER_DPI, doc=None):
    close_doc = False
    if doc is None:
        doc = fitz.open(pdf_path)
        close_doc = True
    try:
        if doc.page_count <= page_index:
            raise ValueError("PDF页数不足：%s 需要页=%d 实际=%d" % (pdf_path, page_index + 1, doc.page_count))
        page = doc.load_page(page_index)
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    finally:
        if close_doc:
            doc.close()


def pdf_rect_to_px_bbox(rect_pt, page_rect_pt, img_w, img_h):
    if rect_pt is None:
        return None
    if page_rect_pt.width <= 0 or page_rect_pt.height <= 0:
        return None
    sx = img_w / float(page_rect_pt.width)
    sy = img_h / float(page_rect_pt.height)

    x0 = (rect_pt.x0 - page_rect_pt.x0) * sx
    y0 = (rect_pt.y0 - page_rect_pt.y0) * sy
    x1 = (rect_pt.x1 - page_rect_pt.x0) * sx
    y1 = (rect_pt.y1 - page_rect_pt.y0) * sy
    return clamp_bbox(x0, y0, x1, y1, img_w, img_h)


def mask_text_regions_on_pil(pil_img, text_boxes_pt, page_rect_pt, pad_mm=TEXT_MASK_PAD_MM):
    if not text_boxes_pt:
        return pil_img
    arr = np.array(pil_img.convert("RGB")).copy()
    H, W = arr.shape[:2]

    sx = W / float(page_rect_pt.width) if page_rect_pt.width > 0 else 1.0
    sy = H / float(page_rect_pt.height) if page_rect_pt.height > 0 else 1.0

    pad_px_x = max(1, int(round((pad_mm * 72.0 / 25.4) * sx)))
    pad_px_y = max(1, int(round((pad_mm * 72.0 / 25.4) * sy)))

    for r in text_boxes_pt:
        x0 = int(round((r.x0 - page_rect_pt.x0) * sx)) - pad_px_x
        y0 = int(round((r.y0 - page_rect_pt.y0) * sy)) - pad_px_y
        x1 = int(round((r.x1 - page_rect_pt.x0) * sx)) + pad_px_x
        y1 = int(round((r.y1 - page_rect_pt.y0) * sy)) + pad_px_y
        x0 = max(0, min(W - 1, x0))
        y0 = max(0, min(H - 1, y0))
        x1 = max(x0 + 1, min(W, x1))
        y1 = max(y0 + 1, min(H, y1))
        arr[y0:y1, x0:x1, :] = 255
    return Image.fromarray(arr)


def _bbox_from_mask(mask, W, H):
    ys, xs = np.where(mask > 0)
    if xs.size < 60 or ys.size < 60:
        return None
    x0, x1 = int(xs.min()), int(xs.max()) + 1
    y0, y1 = int(ys.min()), int(ys.max()) + 1
    if x1 - x0 < 8 or y1 - y0 < 8:
        return None
    if (x1 - x0) >= 0.995 * W and (y1 - y0) >= 0.995 * H:
        return None
    return (x0, y0, x1, y1)


def _reject_border_bbox(x, y, w, h, W, H):
    edge_pad = 8
    touches_edge = (x <= edge_pad or y <= edge_pad or (x + w) >= (W - edge_pad) or (y + h) >= (H - edge_pad))
    huge = (w >= 0.985 * W and h >= 0.985 * H)
    return touches_edge and huge


def _bbox_from_contours_union(mask, W, H):
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    bboxes = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w < max(8, int(W * 0.005)) or h < max(8, int(H * 0.005)):
            continue
        if _reject_border_bbox(x, y, w, h, W, H):
            continue
        bboxes.append((w * h, x, y, w, h))

    if not bboxes:
        return None

    bboxes.sort(key=lambda t: t[0], reverse=True)
    max_area = float(bboxes[0][0])

    keep = []
    for area, x, y, w, h in bboxes[:50]:
        if float(area) >= 0.08 * max_area:
            keep.append((x, y, x + w, y + h))

    if not keep:
        _, x, y, w, h = bboxes[0]
        return (x, y, x + w, y + h)

    x0 = min(k[0] for k in keep)
    y0 = min(k[1] for k in keep)
    x1 = max(k[2] for k in keep)
    y1 = max(k[3] for k in keep)
    return (int(x0), int(y0), int(x1), int(y1))


def find_outer_bbox(pil_img):
    arr = np.array(pil_img.convert("RGB"))
    H, W = arr.shape[0], arr.shape[1]

    if not CV2_OK:
        gray = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]).astype(np.uint8)
        for thr in [250, 245, 240, 235]:
            mask = (gray < thr).astype(np.uint8) * 255
            bbox = _bbox_from_mask(mask, W, H)
            if bbox:
                return bbox
        return (0, 0, W, H)

    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    candidates = []

    try:
        th = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 35, 5)
        th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8), iterations=2)
        th = cv2.dilate(th, np.ones((3, 3), np.uint8), iterations=1)
        b = _bbox_from_contours_union(th, W, H)
        if b:
            candidates.append(b)
    except Exception:
        pass

    try:
        _, th = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8), iterations=2)
        th = cv2.dilate(th, np.ones((3, 3), np.uint8), iterations=1)
        b = _bbox_from_contours_union(th, W, H)
        if b:
            candidates.append(b)
    except Exception:
        pass

    try:
        ed = cv2.Canny(blur, 50, 150)
        ed = cv2.dilate(ed, np.ones((3, 3), np.uint8), iterations=2)
        ed = cv2.morphologyEx(ed, cv2.MORPH_CLOSE, np.ones((9, 9), np.uint8), iterations=1)
        b = _bbox_from_contours_union(ed, W, H)
        if b:
            candidates.append(b)
    except Exception:
        pass

    if not candidates:
        return (0, 0, W, H)

    best = None
    best_area = -1
    for (x0, y0, x1, y1) in candidates:
        bw, bh = x1 - x0, y1 - y0
        if bw <= 0 or bh <= 0:
            continue
        area = bw * bh
        if area > best_area:
            best_area = area
            best = (x0, y0, x1, y1)

    return best if best else candidates[0]


def get_page_bbox_candidates_px(doc, page_index, pil_img):
    info = get_page_struct_info(doc, page_index)
    page_rect = info["page_rect"]
    text_boxes = info["text_boxes"]

    img_w, img_h = pil_img.size
    draw_bbox_px = pdf_rect_to_px_bbox(info["draw_bbox"], page_rect, img_w, img_h)
    image_bbox_px = pdf_rect_to_px_bbox(info["image_bbox"], page_rect, img_w, img_h)

    masked = mask_text_regions_on_pil(pil_img, text_boxes, page_rect, pad_mm=TEXT_MASK_PAD_MM)
    cv_bbox_px = find_outer_bbox(masked)

    return draw_bbox_px, image_bbox_px, cv_bbox_px


def _expand_bbox_clamped(bbox, exp_px, img_w, img_h):
    x0, y0, x1, y1 = bbox
    return clamp_bbox(x0 - exp_px, y0 - exp_px, x1 + exp_px, y1 + exp_px, img_w, img_h)


def make_part_png_bytes_using_ref_bbox(pdf_path, page_index, ref_bbox, ref_size, dpi=RENDER_DPI, doc=None):
    pad_px = mm_to_px(CROP_PAD_MM, dpi)

    img = render_page_to_pil(pdf_path, page_index=page_index, dpi=dpi, doc=doc)
    img_w, img_h = img.size

    draw_px, image_px, cv_px = get_page_bbox_candidates_px(doc, page_index, img)
    bbox_self = union_bbox(union_bbox(draw_px, image_px), cv_px)

    bbox_ref_scaled = None
    if ref_bbox is not None and ref_size is not None:
        rw, rh = ref_size
        if rw > 0 and rh > 0:
            sx = img_w / float(rw)
            sy = img_h / float(rh)
            x0, y0, x1, y1 = ref_bbox
            bbox_ref_scaled = (
                int(round(x0 * sx)),
                int(round(y0 * sy)),
                int(round(x1 * sx)),
                int(round(y1 * sy)),
            )

    bbox = union_bbox(bbox_ref_scaled, bbox_self)
    if bbox is None:
        bbox = (0, 0, img_w, img_h)

    x0, y0, x1, y1 = bbox
    x0, y0, x1, y1 = clamp_bbox(x0 - pad_px, y0 - pad_px, x1 + pad_px, y1 + pad_px, img_w, img_h)

    for _ in range(2):
        crop = img.crop((x0, y0, x1, y1)).convert("RGBA")
        arr = np.array(crop.convert("RGB"))
        gray = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]).astype(np.uint8)
        m = 2
        if (gray[:m, :] < 250).any() or (gray[-m:, :] < 250).any() or (gray[:, :m] < 250).any() or (gray[:, -m:] < 250).any():
            extra_px = mm_to_px(0.6, dpi)
            x0, y0, x1, y1 = _expand_bbox_clamped((x0, y0, x1, y1), extra_px, img_w, img_h)
        else:
            break

    bio = BytesIO()
    img.crop((x0, y0, x1, y1)).convert("RGBA").save(bio, format="PNG", optimize=True)
    return bio.getvalue()


# =========================
# 整拼（单类型）快速算法
# =========================
def _zone_caps_for_w(w):
    left_w = max(0.0, SINGLE_SPLIT_X - SINGLE_MARGIN_L)  # 315
    right_w = max(0.0, (SINGLE_W - SINGLE_MARGIN_R) - (SINGLE_SPLIT_X + SINGLE_SPLIT_GAP_W))  # 265
    left_cap = int(math.floor(left_w / float(w))) if w > 0 else 0
    right_cap = int(math.floor(right_w / float(w))) if w > 0 else 0
    return max(0, left_cap), max(0, right_cap)

def _compute_x_starts_for_w(w):
    """返回每行每列图形的 x 起点列表；块间距固定 10mm；跨 320 时插入间距"""
    W_limit = float(SINGLE_W) - float(SINGLE_MARGIN_R)  # 595
    boundary = float(SINGLE_SPLIT_X)                    # 320
    gap = float(SINGLE_SPLIT_GAP_W)                     # 10

    x = float(SINGLE_MARGIN_L)                          # 5
    starts = []
    guard = 0

    while True:
        # 能放进当前块（不跨 boundary / 不超右边界）
        if (x + w) <= min(boundary, W_limit) + 1e-9:
            starts.append(float(x))
            x += float(w)
        else:
            # 放不下且已经到右边界：结束
            if (x + w) > W_limit + 1e-9:
                break
            # 跨过 boundary：插入 10mm 间距，并进入下一块（boundary += 320）
            x += gap
            boundary += float(SINGLE_SPLIT_X)

            if x > W_limit + 1e-9:
                break

        guard += 1
        if guard > 20000:
            break

    return starts


def _compute_y_starts_for_h(h):
    """返回每列每行图形的 y 起点列表；块间距固定 10mm；跨 464 时插入间距"""
    H_limit = float(SINGLE_H_MAX)
    boundary = 464.0
    gap = 10.0

    y = float(QR_BAND)   # 10
    starts = []
    guard = 0

    while True:
        if (y + h) <= min(boundary, H_limit) + 1e-9:
            starts.append(float(y))
            y += float(h)
        else:
            if (y + h) > H_limit + 1e-9:
                break
            y += gap
            boundary += 464.0
            if y > H_limit + 1e-9:
                break

        guard += 1
        if guard > 20000:
            break

    return starts



def solve_single_type_fixed_width(outer_w, outer_h, need_count):
    """
    ✅ 改成“排布驱动”的求解：
    - 每行列数 = len(_compute_x_starts_for_w(w))
    - 每页最大行数 = len(_compute_y_starts_for_h(h))
    - 这样图形边界天然对齐刀线序列（305->315 这种）
    """
    best = None
    for ori in (0, 1):
        w = float(outer_w if ori == 0 else outer_h)
        h = float(outer_h if ori == 0 else outer_w)
        if w <= 0 or h <= 0:
            continue

        # 每行可放的 x 起点（决定列数）
        x_starts = _compute_x_starts_for_w(w)
        s = len(x_starts)
        if s <= 0:
            continue

        # 每页可放的 y 起点（决定最大行数）
        y_starts = _compute_y_starts_for_h(h)
        rows_per_sheet_max = len(y_starts)
        if rows_per_sheet_max <= 0:
            continue

        r_total = int(math.ceil(float(need_count) / float(s)))

        # 右侧空白（按最后一块实际结束位置算）
        W_limit = float(SINGLE_W) - float(SINGLE_MARGIN_R)
        last_end = float(x_starts[-1]) + float(w)
        right_blank = float(max(0.0, W_limit - last_end))

        cand = {
            "ori": ori,
            "w": w,
            "h": h,
            "k_cols": int(s),
            "rows_total": int(r_total),
            "rows_per_sheet_max": int(rows_per_sheet_max),
            "right_blank": float(right_blank),
            "x_starts": x_starts,
            "y_starts": y_starts,
        }

        if best is None:
            best = cand
        else:
            if cand["right_blank"] < best["right_blank"] - 1e-9:
                best = cand
            elif abs(cand["right_blank"] - best["right_blank"]) <= 1e-9:
                if cand["k_cols"] > best["k_cols"]:
                    best = cand

    return best

def build_single_placements_full_rows(best_base, rows_seg):
    """
    ✅ 用 x_starts/y_starts 直接生成排布：
    - 块间距就是 10mm（因为 x_starts/y_starts 已经插入了 gap）
    - 图形边界会对齐刀线序列（例如 245-305, gap 305-315, 再 315-375）
    """
    w = float(best_base["w"])
    h = float(best_base["h"])
    s = int(best_base["k_cols"])
    rotated = (int(best_base["ori"]) == 1)

    x_starts = list(best_base.get("x_starts") or [])
    y_starts_all = list(best_base.get("y_starts") or [])

    # 只取本段需要的行数
    y_starts = y_starts_all[:int(rows_seg)]
    if len(x_starts) != s:
        # 保险：若外部改了 k_cols，强制对齐
        s = len(x_starts)

    placements = []
    for rr in range(len(y_starts)):
        y = float(y_starts[rr])
        for cc in range(s):
            x = float(x_starts[cc])
            placements.append({
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "rot": bool(rotated),
            })

    # 段高度：以最后一行结束位置为准（含块间距）
    if y_starts:
        H_seg = float(y_starts[-1]) + float(h)
    else:
        H_seg = float(QR_BAND)

    seg_best = dict(best_base)
    seg_best["rows_per_sheet"] = int(rows_seg)
    seg_best["items_per_sheet"] = int(rows_seg) * int(s)
    seg_best["H"] = float(H_seg)
    return placements, seg_best

def draw_label_in_band(page, text, x_mm, y_mm, w_mm):
    fontfile = pick_cjk_fontfile()
    xpt = mm_to_pt(x_mm + 1.0)
    ypt = mm_to_pt(y_mm - 1.5)
    max_w = mm_to_pt(w_mm - 2.0)
    txt = trim_text_to_width(text, LABEL_FONT_SIZE, max_w)
    if fontfile:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=LABEL_FONT_SIZE, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=LABEL_FONT_SIZE, fontname="helv", color=(0, 0, 0))


def _draw_single_segment_on_page_no_cuts(page, seg, img_bytes, yoff_mm):
    placements = seg["placements"]
    label_text = seg["label_text"]
    qr_bytes = seg["qr_bytes"]

    Wpt = mm_to_pt(SINGLE_W)
    yoff_pt = mm_to_pt(yoff_mm)

    # 右上QR
    qr_rect = fitz.Rect(Wpt - mm_to_pt(QR_W), yoff_pt, Wpt, yoff_pt + mm_to_pt(QR_H))
    if qr_bytes:
        page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)

    # 标签
    if label_text and placements:
        first = placements[0]
        draw_label_in_band(page, label_text, first["x"], yoff_mm + first["y"], first["w"])

    # 贴图
    for p in placements:
        x, y, w, h = p["x"], (yoff_mm + p["y"]), p["w"], p["h"]
        outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
        if DRAW_PART_OUTER_BOX:
            page.draw_rect(outer, color=(0, 0, 0), width=0.5)

        inner = fitz.Rect(
            outer.x0 + mm_to_pt(INNER_MARGIN_MM),
            outer.y0 + mm_to_pt(INNER_MARGIN_MM),
            outer.x1 - mm_to_pt(INNER_MARGIN_MM),
            outer.y1 - mm_to_pt(INNER_MARGIN_MM),
        )
        if img_bytes:
            page.insert_image(inner, stream=img_bytes, keep_proportion=True,
                              rotate=(90 if p["rot"] else 0))


def _pack_segments_ffd_to_fixed_sheets(segments, max_h_mm=1500.0, gap_mm=0.0):
    segs = list(segments)
    segs.sort(key=lambda s: (-float(s["best"]["H"]), str(s.get("type_key", "")), str(s.get("label_text", ""))))

    sheets = []
    for seg in segs:
        h = float(seg["best"]["H"])
        placed = False
        for sh in sheets:
            add = h + (gap_mm if sh["segments"] else 0.0)
            if sh["used_h"] + add <= float(max_h_mm) + 1e-9:
                if sh["segments"]:
                    sh["used_h"] += float(gap_mm)
                sh["segments"].append(seg)
                sh["used_h"] += h
                placed = True
                break
        if not placed:
            sheets.append({"segments": [seg], "used_h": h})
    return sheets

def _render_one_single_sheet_doc(sheet, is_contour, fixed_h_mm=SINGLE_H_MAX):
    W = float(SINGLE_W)
    H = float(fixed_h_mm)

    doc = fitz.open()
    page = doc.new_page(width=mm_to_pt(W), height=mm_to_pt(H))

    # 外框
    page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=(0, 0, 0), width=1.0)

    yoff = 0.0

    for idx, seg in enumerate(sheet["segments"]):
        img_bytes = seg["img_cont"] if is_contour else seg["img_body"]

        # ✅ 段与段之间：横刀线 + 左侧刻度（不再依赖 type_key，直接补齐蓝圈缺失）
        if idx > 0:
            _draw_horizontal_cutline(page, yoff, lw=1.0)
            _draw_left_edge_tick(page, yoff, tick_len_mm=MARK_LEN, lw=1.0)

        # 先贴图
        _draw_single_segment_on_page_no_cuts(page, seg, img_bytes, yoff_mm=yoff)

        # ✅ 再画“该段自己的刀线刻度”（修复红圈乱标 + 蓝圈缺失）
        _draw_segment_cut_ticks(page, seg, yoff_mm=yoff, page_h_mm=H)

        yoff += float(seg["best"]["H"])

    return doc

# =========================
# safe_save
# =========================
def safe_save(doc, out_path):
    ensure_dir(os.path.dirname(out_path))
    tmp_pdf = os.path.join(
        os.path.dirname(out_path),
        os.path.splitext(os.path.basename(out_path))[0] + "_tmp.pdf"
    )

    try:
        if os.path.exists(tmp_pdf):
            os.remove(tmp_pdf)
    except Exception:
        pass

    doc.save(tmp_pdf, garbage=4, deflate=True, incremental=False)
    doc.close()

    try:
        os.replace(tmp_pdf, out_path)
        return out_path
    except PermissionError:
        ts = time.strftime("%Y%m%d_%H%M%S")
        alt_pdf = os.path.join(
            os.path.dirname(out_path),
            os.path.splitext(os.path.basename(out_path))[0] + "_%s.pdf" % ts
        )
        try:
            shutil.move(tmp_pdf, alt_pdf)
        except Exception:
            shutil.copyfile(tmp_pdf, alt_pdf)
            try:
                os.remove(tmp_pdf)
            except Exception:
                pass
        return alt_pdf


# =========================
# 主入口：run()
# =========================
def run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None):
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2

    if cfg:
        DEST_DIR = cfg.get("DEST_DIR", DEST_DIR)
        IN_PDF_ARCHIVE_DIR = cfg.get("TEST_DIR", IN_PDF_ARCHIVE_DIR)
        DEST_DIR1 = cfg.get("DEST_DIR1", DEST_DIR1)
        DEST_DIR2 = cfg.get("DEST_DIR2", DEST_DIR2)

    ensure_dir(IN_PDF_ARCHIVE_DIR)
    ensure_dir(DEST_DIR1)
    ensure_dir(DEST_DIR2)

    # 收集输入pdf并归档
    if input_pdfs is not None:
        archived_paths = list(input_pdfs)
    else:
        pdfs = []
        for fn in os.listdir(DEST_DIR):
            if fn.lower().endswith(".pdf"):
                pdfs.append(os.path.join(DEST_DIR, fn))
        pdfs.sort()

        archived_paths = []
        for src in pdfs:
            try:
                dst = archive_input_pdf_to_dir(src, IN_PDF_ARCHIVE_DIR)
                archived_paths.append(dst)
            except Exception as e:
                _log(log_cb, "⚠️ 复制到test失败：%s err=%s" % (os.path.basename(src), repr(e)))

    if not archived_paths:
        raise RuntimeError("没有可处理PDF（archived_paths为空）")

    # 整拼段（用于 single_i）
    global_single_segments = []

    ok_list = []
    skip_list = []

    N_total = len(archived_paths)
    for idx, path in enumerate(archived_paths, start=1):
        type_name_raw = os.path.splitext(os.path.basename(path))[0]
        if progress_cb:
            progress_cb(idx - 1, N_total, "处理中 %d/%d  %s" % (idx - 1, N_total, type_name_raw))

        d = None
        try:
            d = fitz.open(path)
            pc = d.page_count
            if pc < 2:
                raise RuntimeError("PDF页数不足2页")

            pair_count = pc // 2
            if pair_count <= 0:
                raise RuntimeError("PDF无有效页对(页数=%d)" % pc)

            A, B, N = parse_A_B_N_from_filename(path)

            # ✅ N < 10：不拼版，直接跳过
            if int(N) < 10:
                skip_list.append((type_name_raw, "N<10_skip_no_layout"))
                _log(log_cb, "⏭️ N<10 不拼版跳过：%s  N=%d" % (type_name_raw, int(N)))
                continue

            outer_w = int(round(A + 2.0 * OUTER_EXT))
            outer_h = int(round(B + 2.0 * OUTER_EXT))

            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)

            best_base = solve_single_type_fixed_width(outer_w, outer_h, N)
            if best_base is None:
                skip_list.append((type_name_raw, "single_no_solution_need>=10"))
                _log(log_cb, "⚠️ 整拼无解 -> SKIP: %s" % type_name_raw)
                continue

            for pi in range(pair_count):
                page_body = 2 * pi
                page_cont = 2 * pi + 1

                # ref bbox
                ref_img = render_page_to_pil(path, page_index=page_cont, dpi=RENDER_DPI, doc=d)
                draw_px, img_px, cv_px = get_page_bbox_candidates_px(d, page_cont, ref_img)
                if draw_px is not None:
                    ref_bbox = union_bbox(draw_px, cv_px)
                elif img_px is not None:
                    ref_bbox = union_bbox(img_px, cv_px)
                else:
                    ref_bbox = cv_px
                ref_size = ref_img.size

                img_body = make_part_png_bytes_using_ref_bbox(path, page_body, ref_bbox, ref_size, dpi=RENDER_DPI, doc=d)
                img_cont = make_part_png_bytes_using_ref_bbox(path, page_cont, ref_bbox, ref_size, dpi=RENDER_DPI, doc=d)

                tid = "%s@P%d" % (type_name_raw, pi + 1)

                # N>=10：整拼
                s = int(best_base["k_cols"])
                r_total = int(best_base["rows_total"])
                print_total = r_total * s
                gift = print_total - int(N)

                max_rows = int(best_base["rows_per_sheet_max"])

                remain_rows = r_total
                seg_cnt = 0
                qr_bytes = make_qr_png_bytes(qr_text)

                while remain_rows > 0:
                    seg_cnt += 1
                    rows_seg = min(max_rows, remain_rows)
                    placements, seg_best = build_single_placements_full_rows(best_base, rows_seg)

                    global_single_segments.append({
                        "type_key": tid,
                        "tid": tid,
                        "best": seg_best,
                        "placements": placements,
                        "qr_bytes": qr_bytes,
                        "label_text": label_text,
                        "img_body": img_body,
                        "img_cont": img_cont,
                    })

                    remain_rows -= rows_seg

                ok_list.append((tid, seg_cnt))
                _log(log_cb,
                     "✅ 整拼OK: %s seg=%d  s=%d  r=%d  print=%d(送=%d)  right_blank=%.1fmm"
                     % (tid, seg_cnt, s, r_total, print_total, gift, float(best_base["right_blank"])))

        except Exception as e:
            skip_list.append((type_name_raw, repr(e)))
            _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, repr(e)))
        finally:
            try:
                if d is not None:
                    d.close()
            except Exception:
                pass

    # =========================
    # 输出 single_i：把所有整拼段合在一起，尽量塞满1500
    # =========================
    single_outputs_p1 = []
    single_outputs_p2 = []

    if global_single_segments:
        packed = _pack_segments_ffd_to_fixed_sheets(global_single_segments, max_h_mm=SINGLE_H_MAX, gap_mm=0.0)

        single_counter = 0
        for sh in packed:
            single_counter += 1
            name = "single_%d.pdf" % single_counter
            out_path1 = os.path.join(DEST_DIR1, name)
            out_path2 = os.path.join(DEST_DIR2, name)

            doc1 = _render_one_single_sheet_doc(sh, is_contour=False, fixed_h_mm=SINGLE_H_MAX)
            doc2 = _render_one_single_sheet_doc(sh, is_contour=True, fixed_h_mm=SINGLE_H_MAX)

            p1 = safe_save(doc1, out_path1)
            p2 = safe_save(doc2, out_path2)
            single_outputs_p1.append(p1)
            single_outputs_p2.append(p2)
            _log(log_cb, "📄 %s：segments=%d used_h=%.1f/%d  %s | %s"
                 % (name, len(sh["segments"]), sh["used_h"], int(SINGLE_H_MAX), p1, p2))

    if progress_cb:
        progress_cb(N_total, N_total, "完成 %d/%d" % (N_total, N_total))


    return {
        "single_p1_files": single_outputs_p1,
        "single_p2_files": single_outputs_p2,
        "ok": ok_list,
        "skip": skip_list,
    }


def main():
    res = run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None)
    print("DONE:")
    print("  single_p1_files:", len(res.get("single_p1_files") or []))
    print("  single_p2_files:", len(res.get("single_p2_files") or []))


if __name__ == "__main__":
    main()