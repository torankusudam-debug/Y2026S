import os
import re
import math
import time
import shutil
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
DEST_DIR = r"D:\test_data\dest"
IN_PDF_ARCHIVE_DIR = r"D:\test_data\test"
DEST_DIR1 = r"D:\test_data\gest"
DEST_DIR2 = r"D:\test_data\pest"
OUTER_EXT = 2.0                             # 刀线外扩（mm），避免裁剪时漏掉边缘内容
INNER_MARGIN_MM = 3.0                       # 内边距（mm），避免刀线太靠近内容，影响美观和实用, 也避免裁剪时误伤内容。
MARK_LEN = 5.0                              # 刀线/刻度长度（mm）
SINGLE_W = 600.0
SINGLE_H_MAX = 1500.0
SINGLE_MARGIN_L = 5.0                       # 单拼段左边距（mm），避免刀线太靠近边缘，影响美观和实用，也避免裁剪时误伤内容。
SINGLE_MARGIN_R = 5.0                       # 单拼段右边距（mm），避免刀线太靠近边缘，影响美观和实用，也避免裁剪时误伤内容。
SINGLE_SPLIT_X = 320.0
SINGLE_SPLIT_GAP_W = 10.0                   # 单拼段间隔（mm），避免刀线太靠近内容，影响美观和实用，也避免裁剪时误伤内容，同时也能让用户更清晰地看到分段。
SINGLE_RESERVED_W = SINGLE_SPLIT_GAP_W + (SINGLE_MARGIN_L + SINGLE_MARGIN_R)  # 20
SINGLE_USABLE_W = SINGLE_W - SINGLE_RESERVED_W  # 580
QR_BAND = 10.0                              # QR 区高度（mm），从单拼段顶部开始算起, 这个区域内不画刀线刻度，且优先放 QR 码（如果有的话），避免和内容重叠。
QR_W = 10.0                                 # QR 码宽度（mm），放在单拼段右上角，和右边距一起占满右侧区域，避免和内容重叠。
QR_H = 10.0                                 # QR 码高度（mm），放在单拼段右上角，和 QR_BAND 一起占满顶部区域，避免和内容重叠。
LABEL_FONT_SIZE = 10                        # 标签字体大小（pt），放在单拼段右上角，和 QR 码一起占满右侧区域，避免和内容重叠。
RENDER_DPI = 600                            # PDF渲染分辨率（DPI），600 DPI 可以让毫米级的细节在图像上有足够的像素表现，便于后续的分析和裁剪。
CROP_PAD_MM = 1.2                           # 裁剪边界额外扩展（mm），避免裁剪时漏掉边缘内容，同时也能让成品看起来更自然一些。
TEXT_MASK_PAD_MM = 1.2
DRAW_PART_OUTER_BOX = False
CELL_GAP_MM = 10.0
MARGIN_LR_MM = 5.0
MARGIN_TOP_MM = 10.0
MARGIN_BOTTOM_MM = 5.0
SAFE_PAD_MM = 3.0
PAGE_W = 600.0
PAGE_H_MAX = 1500.0
PAGE_MARGIN_L = MARGIN_LR_MM
PAGE_MARGIN_R = MARGIN_LR_MM
SPLIT_X = 320.0
SPLIT_GAP_W = 10.0
MIX_ENABLE = True
MIX_THRESHOLD = 10
LEFT_BLOCK_W = float(SPLIT_X - PAGE_MARGIN_L)
RIGHT_BLOCK_W = float((PAGE_W - PAGE_MARGIN_R) - (SPLIT_X + SPLIT_GAP_W))
TOTAL_USABLE_W = float((PAGE_W - PAGE_MARGIN_R) - PAGE_MARGIN_L)
def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)
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
    w = float(pls[0].get("w", 0.0))
    h = float(pls[0].get("h", 0.0))
    if w <= 0 or h <= 0:
        return
    xs = sorted({round(float(p["x"]), 3) for p in pls})
    ys = sorted({round(float(p["y"]), 3) for p in pls})
    x_lines = set(xs)
    for x in xs:
        x_lines.add(round(x + w, 3))
    x_lines = sorted(x_lines)
    if not x_lines:
        return
    W = float(SINGLE_W)
    x_min = float(min(x_lines))
    x_max = float(max(x_lines))
    for x in x_lines:
        x_draw = float(x)
        if abs(x - x_min) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        elif abs(x - x_max) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        if 0.0 - 1e-6 <= x_draw <= W + 1e-6:
            _draw_top_tick_at_y(page, x_draw, yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)
    y_lines = set(ys)
    for y in ys:
        y_lines.add(round(y + h, 3))
    y_lines = sorted(y_lines)
    if not y_lines:
        return
    y_min = float(min(y_lines))
    y_max = float(max(y_lines))
    for y in y_lines:
        Y = float(yoff_mm) + float(y)
        if abs(y - y_max) <= 1e-6:
            Y = _snap_edge_mm(Y, 0.0, float(page_h_mm), EDGE_SNAP_MM)
        if 0.0 - 1e-6 <= Y <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, Y, tick_len_mm=MARK_LEN, lw=1.0)
def _draw_horizontal_cutline(page, y_mm, lw=1.0):
    ypt = mm_to_pt(float(y_mm))
    Wpt = mm_to_pt(float(SINGLE_W))
    page.draw_line(fitz.Point(0, ypt), fitz.Point(Wpt, ypt), color=(0, 0, 0), width=lw)
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
def _draw_label_top_left(page, text, x_left_mm, y_top_mm, max_w_mm):
    fontfile = pick_cjk_fontfile()
    fontsize = LABEL_FONT_SIZE
    max_w_pt = mm_to_pt(max(5.0, float(max_w_mm)))
    txt = trim_text_to_width(text, fontsize, max_w_pt)
    xpt = mm_to_pt(float(x_left_mm) + 1.0)
    ypt = mm_to_pt(float(y_top_mm) + 7.5)
    if fontfile:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontname="helv", color=(0, 0, 0))
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
def _compute_x_starts_for_w(w):
    """返回每行每列图形的 x 起点列表；块间距固定 10mm；跨 320 时插入间距"""
    W_limit = float(SINGLE_W) - float(SINGLE_MARGIN_R)  # 595
    boundary = float(SINGLE_SPLIT_X)                    # 320
    gap = float(SINGLE_SPLIT_GAP_W)                     # 10
    x = float(SINGLE_MARGIN_L)                          # 5
    starts = []
    guard = 0
    while True:
        if (x + w) <= min(boundary, W_limit) + 1e-9:
            starts.append(float(x))
            x += float(w)
        else:
            if (x + w) > W_limit + 1e-9:
                break
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
        x_starts = _compute_x_starts_for_w(w)
        s = len(x_starts)
        if s <= 0:
            continue
        y_starts = _compute_y_starts_for_h(h)
        rows_per_sheet_max = len(y_starts)
        if rows_per_sheet_max <= 0:
            continue
        r_total = int(math.ceil(float(need_count) / float(s)))
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
    y_starts = y_starts_all[:int(rows_seg)]
    if len(x_starts) != s:
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
    qr_rect = fitz.Rect(Wpt - mm_to_pt(QR_W), yoff_pt, Wpt, yoff_pt + mm_to_pt(QR_H))
    if qr_bytes:
        page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)
    if label_text and placements:
        first = placements[0]
        draw_label_in_band(page, label_text, first["x"], yoff_mm + first["y"], first["w"])
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
    page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=(0, 0, 0), width=1.0)
    yoff = 0.0
    for idx, seg in enumerate(sheet["segments"]):
        img_bytes = seg["img_cont"] if is_contour else seg["img_body"]
        if idx > 0:
            _draw_horizontal_cutline(page, yoff, lw=1.0)
            _draw_left_edge_tick(page, yoff, tick_len_mm=MARK_LEN, lw=1.0)
        _draw_single_segment_on_page_no_cuts(page, seg, img_bytes, yoff_mm=yoff)
        _draw_segment_cut_ticks(page, seg, yoff_mm=yoff, page_h_mm=H)
        yoff += float(seg["best"]["H"])
    return doc
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
def _collect_archived_paths(input_pdfs=None, log_cb=None):
    if input_pdfs is not None:
        paths = list(input_pdfs)
        if not paths:
            raise RuntimeError("没有可处理PDF（input_pdfs为空）")
        return paths
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
    return archived_paths
def _open_pdf_job(path):
    """提取两套流程共用的 PDF 基础信息。"""
    type_name_raw = os.path.splitext(os.path.basename(path))[0]
    doc = fitz.open(path)
    pc = doc.page_count
    if pc < 2:
        raise RuntimeError("PDF页数不足2页")
    pair_count = pc // 2
    if pair_count <= 0:
        raise RuntimeError("PDF无有效页对(页数=%d)" % pc)
    A, B, N = parse_A_B_N_from_filename(path)
    return type_name_raw, doc, pair_count, A, B, N
def _build_pair_images(path, doc, pair_index):
    """
    使用轮廓页提取参考 bbox，再裁出图形页与轮廓页 PNG。
    保持与原流程一致，只做去重封装。
    """
    page_body = 2 * pair_index
    page_cont = 2 * pair_index + 1
    ref_img = render_page_to_pil(path, page_index=page_cont, dpi=RENDER_DPI, doc=doc)
    draw_px, img_px, cv_px = get_page_bbox_candidates_px(doc, page_cont, ref_img)
    if draw_px is not None:
        ref_bbox = union_bbox(draw_px, cv_px)
    elif img_px is not None:
        ref_bbox = union_bbox(img_px, cv_px)
    else:
        ref_bbox = cv_px
    ref_size = ref_img.size
    img_body = make_part_png_bytes_using_ref_bbox(path, page_body, ref_bbox, ref_size, dpi=RENDER_DPI, doc=doc)
    img_cont = make_part_png_bytes_using_ref_bbox(path, page_cont, ref_bbox, ref_size, dpi=RENDER_DPI, doc=doc)
    return img_body, img_cont
def _save_dual_docs(doc1, doc2, out_path1, out_path2):
    p1 = safe_save(doc1, out_path1)
    p2 = safe_save(doc2, out_path2)
    return p1, p2
def _append_skip(skip_list, log_cb, type_name_raw, reason):
    skip_list.append((type_name_raw, reason))
    _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, reason))
def _finish_stage_progress(progress_cb, total):
    if progress_cb:
        progress_cb(total, total, "完成 %d/%d" % (total, total))
def _iter_open_pdf_jobs(archived_paths, progress_cb=None, on_error=None):
    """
    统一处理：进度上报 + 打开PDF + 基础信息解析 + 关闭PDF。
    """
    n_total = len(archived_paths)
    for idx, path in enumerate(archived_paths, start=1):
        type_name_raw = os.path.splitext(os.path.basename(path))[0]
        if progress_cb:
            progress_cb(idx - 1, n_total, "处理中 %d/%d  %s" % (idx - 1, n_total, type_name_raw))
        doc = None
        try:
            type_name_raw, doc, pair_count, A, B, N = _open_pdf_job(path)
            yield {
                "path": path,
                "type_name_raw": type_name_raw,
                "doc": doc,
                "pair_count": pair_count,
                "A": A,
                "B": B,
                "N": N,
            }
        except Exception as e:
            if on_error:
                on_error(type_name_raw, e)
        finally:
            try:
                if doc is not None:
                    doc.close()
            except Exception:
                pass
def _iter_pair_payloads(path, doc, type_name_raw, pair_count):
    """统一页对遍历：给出 tid + 图形/轮廓两张裁图。"""
    for pi in range(pair_count):
        img_body, img_cont = _build_pair_images(path, doc, pi)
        tid = "%s@P%d" % (type_name_raw, pi + 1)
        yield tid, img_body, img_cont
def _run_single_stage(archived_paths, progress_cb=None, log_cb=None):
    """
    一次遍历完成：
    1) 整拼评估 + 输出素材收集（N>=MIX_THRESHOLD）
    2) 混拼评估 + MIX_POOL 构建（N<MIX_THRESHOLD）
    这样 stage2 只消费 mix_pool，不再重读/重评估 PDF。
    """
    global_single_segments = []
    mix_pool = []
    ok_single = []
    skip_single = []
    ok_mix_pool = []
    skip_mix_pool = []
    n_total = len(archived_paths)
    def _on_open_error(type_name_raw, e):
        reason = repr(e)
        _append_skip(skip_single, log_cb, type_name_raw, reason)
        skip_mix_pool.append((type_name_raw, reason))
    for job in _iter_open_pdf_jobs(archived_paths, progress_cb=progress_cb, on_error=_on_open_error):
        path = job["path"]
        type_name_raw = job["type_name_raw"]
        d = job["doc"]
        pair_count = int(job["pair_count"])
        A = float(job["A"])
        B = float(job["B"])
        N = int(job["N"])
        try:
            outer_w0 = float(A + 2.0 * OUTER_EXT)
            outer_h0 = float(B + 2.0 * OUTER_EXT)
            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)
            qr_bytes = make_qr_png_bytes(qr_text)
            if N < MIX_THRESHOLD:
                skip_single.append((type_name_raw, "N<%d_skip_no_layout" % MIX_THRESHOLD))
                _log(log_cb, "⏭️ N<%d 不拼版跳过：%s  N=%d" % (MIX_THRESHOLD, type_name_raw, N))
                Mi, Ni, rot90 = _mix_choose_orient(outer_w0, outer_h0)
                for tid, img_body, img_cont in _iter_pair_payloads(path, d, type_name_raw, pair_count):
                    mix_pool.append({
                        "tid": tid,
                        "label_text": label_text,
                        "qr_bytes": qr_bytes,
                        "img_body": img_body,
                        "img_cont": img_cont,
                        "rem": int(N),
                        "W": float(Mi),
                        "H": float(Ni),
                        "rot90": bool(rot90),
                    })
                    ok_mix_pool.append((tid, "mix_pool"))
                    _log(log_cb, "🧩 MIX_POOL: %s N=%d Mi=%.1f Ni=%.1f rot90=%s"
                         % (tid, int(N), float(Mi), float(Ni), str(bool(rot90))))
                continue
            skip_mix_pool.append((type_name_raw, "skip_N_ge_%d" % MIX_THRESHOLD))
            outer_w = int(round(outer_w0))
            outer_h = int(round(outer_h0))
            best_base = solve_single_type_fixed_width(outer_w, outer_h, N)
            if best_base is None:
                skip_single.append((type_name_raw, "single_no_solution_need>=10"))
                _log(log_cb, "⚠️ 整拼无解 -> SKIP: %s" % type_name_raw)
                continue
            s = int(best_base["k_cols"])
            r_total = int(best_base["rows_total"])
            print_total = r_total * s
            gift = print_total - int(N)
            max_rows = int(best_base["rows_per_sheet_max"])
            for tid, img_body, img_cont in _iter_pair_payloads(path, d, type_name_raw, pair_count):
                remain_rows = r_total
                seg_cnt = 0
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
                ok_single.append((tid, seg_cnt))
                _log(log_cb, "✅ 整拼OK: %s seg=%d  s=%d  r=%d  print=%d(送=%d)  right_blank=%.1fmm"
                     % (tid, seg_cnt, s, r_total, print_total, gift, float(best_base["right_blank"])))
        except Exception as e:
            reason = repr(e)
            _append_skip(skip_single, log_cb, type_name_raw, reason)
            skip_mix_pool.append((type_name_raw, reason))
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
            page_h_mm = float(sh.get("used_h", 0.0))
            if page_h_mm <= 0:
                page_h_mm = float(SINGLE_H_MAX)
            doc1 = _render_one_single_sheet_doc(sh, is_contour=False, fixed_h_mm=page_h_mm)
            doc2 = _render_one_single_sheet_doc(sh, is_contour=True, fixed_h_mm=page_h_mm)
            p1, p2 = _save_dual_docs(doc1, doc2, out_path1, out_path2)
            single_outputs_p1.append(p1)
            single_outputs_p2.append(p2)
            _log(log_cb, "📄 %s：segments=%d used_h=%.1f  %s | %s"
                 % (name, len(sh["segments"]), float(page_h_mm), p1, p2))
    _finish_stage_progress(progress_cb, n_total)
    return {
        "single_p1_files": single_outputs_p1,
        "single_p2_files": single_outputs_p2,
        "ok": ok_single,
        "skip": skip_single,
        "mix_pool": mix_pool,
        "ok_mix_pool": ok_mix_pool,
        "skip_mix_pool": skip_mix_pool,
    }
def _mix_choose_orient(w0, h0):
    w0 = float(w0)
    h0 = float(h0)
    w_a, h_a, r_a = w0, h0, False
    w_b, h_b, r_b = h0, w0, True
    a_ok = (w_a <= LEFT_BLOCK_W + 1e-9)
    b_ok = (w_b <= LEFT_BLOCK_W + 1e-9)
    if a_ok and b_ok:
        return (w_a, h_a, r_a) if h_a >= h_b else (w_b, h_b, r_b)
    if a_ok:
        return (w_a, h_a, r_a)
    if b_ok:
        return (w_b, h_b, r_b)
    return (w_a, h_a, r_a) if w_a <= w_b else (w_b, h_b, r_b)
def _iter_mix_orients(w, h, rot90):
    w = float(w)
    h = float(h)
    rot90 = bool(rot90)
    opts = [(w, h, rot90)]
    if abs(w - h) > 1e-6:
        opts.append((h, w, (not rot90)))
    out = []
    seen = set()
    for wi, hi, ri in opts:
        key = (round(wi, 6), round(hi, 6), bool(ri))
        if key in seen:
            continue
        seen.add(key)
        out.append((float(wi), float(hi), bool(ri)))
    return out
def _advance_mix_y(y, row_used_h):
    y2 = float(y) + float(row_used_h)
    boundary = 464.0
    while boundary <= float(y) + 1e-9:
        boundary += 464.0
    if float(y) < boundary and y2 > boundary + 1e-9:
        y2 += 10.0
    return float(y2)
def _collect_mix_base_heights(types_list, max_w):
    hs = set()
    for t in types_list:
        if int(t.get("rem", 0)) <= 0:
            continue
        for wi, hi, _rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)):
            if float(wi) <= float(max_w) + 1e-9:
                hs.add(float(hi))
    return sorted(hs, reverse=True)
def pack_mix_by_height_rule(mix_types, page_max_h=PAGE_H_MAX):
    X_LEFT0 = float(PAGE_MARGIN_L)
    X_LEFT1 = float(SPLIT_X)
    X_RIGHT0 = float(SPLIT_X + SPLIT_GAP_W)
    X_RIGHT1 = float(PAGE_W - PAGE_MARGIN_R)
    gap = float(CELL_GAP_MM)
    types_list = []
    for t in mix_types:
        tt = dict(t)
        tt["rem"] = int(tt.get("rem", tt.get("N", 0)) or 0)
        types_list.append(tt)
    for t in types_list:
        fit_lane = False
        for wi, _hi, _rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)):
            if float(wi) <= float(LEFT_BLOCK_W) + 1e-9:
                fit_lane = True
                break
        if int(t["rem"]) > 0 and (not fit_lane):
            print("⚠️ MIX_SKIP_TOO_WIDE_FOR_LANE:", t.get("tid"), "minW=%.2f > %.2f" % (min(float(t["W"]), float(t["H"])), float(LEFT_BLOCK_W)))
            t["rem"] = 0
    def _pick_orient_and_layout(t, base_h, rem_w):
        best = None
        best_rank = None
        rem = int(t.get("rem", 0))
        for Wi, Hi, rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)):
            if Hi > base_h + 1e-9:
                continue
            if Wi > rem_w + 1e-9:
                continue
            stack_cap = int(math.floor((base_h + gap) / (Hi + gap))) if Hi > 0 else 1
            if stack_cap <= 0:
                stack_cap = 1
            max_cols = int(math.floor((rem_w + gap) / (Wi + gap))) if Wi > 0 else 0
            if max_cols <= 0:
                continue
            need_cols = int(math.ceil(float(rem) / float(stack_cap))) if rem > 0 else 1
            cols = min(max_cols, need_cols)
            if cols <= 0:
                cols = 1
            place_cnt = min(rem, cols * stack_cap)
            block_w = float(cols) * Wi + float(max(0, cols - 1)) * gap
            leftover = rem_w - block_w
            used_area = float(place_cnt) * Wi * Hi
            denom = max(1.0, float(block_w) * float(base_h))
            density = used_area / denom
            rank = (-density, float(leftover), -float(used_area), -float(place_cnt), float(Hi))
            if best is None or rank < best_rank:
                best = {
                    "w": float(Wi),
                    "h": float(Hi),
                    "rot90": bool(rot90),
                    "stack_cap": int(stack_cap),
                    "cols": int(cols),
                    "block_w": float(block_w),
                    "leftover": float(leftover),
                    "used_area": float(used_area),
                    "place_cnt": int(place_cnt),
                }
                best_rank = rank
        return best
    def _build_one_region_row(work_list, base_h, y, x_start, x_end):
        x = float(x_start)
        row_blocks = []
        row_added = []
        while True:
            rem_w = float(x_end) - float(x)
            if rem_w <= 1e-6:
                break
            cand = None
            cand_plan = None
            cand_rank = None
            for t in work_list:
                if int(t["rem"]) <= 0:
                    continue
                plan = _pick_orient_and_layout(t, base_h, rem_w)
                if plan is None:
                    continue
                rank = (-float(plan["used_area"]) / max(1.0, float(plan["block_w"]) * float(base_h)),
                        float(plan["leftover"]),
                        -float(plan["used_area"]),
                        float(plan["h"]))
                if cand is None or rank < cand_rank:
                    cand = t
                    cand_plan = plan
                    cand_rank = rank
            if cand is None:
                break
            Wi = float(cand_plan["w"])
            Hi = float(cand_plan["h"])
            stack_cap = int(cand_plan["stack_cap"])
            cols = int(cand_plan["cols"])
            x0 = float(x)
            block_w = float(cand_plan["block_w"])
            x1 = x0 + block_w
            block_min_x = None
            block_max_x = None
            block_max_y = None
            for c in range(cols):
                col_x = x0 + float(c) * (Wi + gap)
                for s in range(stack_cap):
                    if int(cand["rem"]) <= 0:
                        break
                    px = col_x
                    py = float(y) + float(QR_BAND) + float(s) * (Hi + gap)
                    row_added.append({
                        "type": cand,
                        "tid": cand["tid"],
                        "x": px,
                        "y": py,
                        "w": Wi,
                        "h": Hi,
                        "rot90": bool(cand_plan["rot90"]),
                    })
                    cand["rem"] -= 1
                    bx0 = float(px)
                    bx1 = float(px) + float(Wi)
                    by1 = float(py) + float(Hi)
                    if block_min_x is None:
                        block_min_x = bx0
                        block_max_x = bx1
                        block_max_y = by1
                    else:
                        block_min_x = min(block_min_x, bx0)
                        block_max_x = max(block_max_x, bx1)
                        block_max_y = max(block_max_y, by1)
            if block_min_x is not None:
                row_blocks.append({
                    "type": cand,
                    "tid": cand["tid"],
                    "x0": float(block_min_x),
                    "x1": float(block_max_x),
                    "y0": float(y),
                    "y1": float(block_max_y),
                })
            x = x1
        if not row_added:
            return [], [], 0.0, 0.0
        row_max_end = 0.0
        for p in row_added:
            row_max_end = max(row_max_end, float(p["y"]) + float(p["h"]))
        row_used_h = max(float(QR_BAND), row_max_end - float(y))
        return row_blocks, row_added, float(row_used_h), float(row_max_end)
    def _find_best_region_row(work_list, region):
        region_w = float(region["x1"] - region["x0"])
        if region_w <= 1e-6 or float(region["y"]) >= float(page_max_h) - 1e-6:
            return None
        hs = _collect_mix_base_heights(work_list, region_w)
        cand_hs = hs[:8] if len(hs) > 8 else hs
        best = None
        best_score = None
        for bh in cand_hs:
            sim_list = [dict(t) for t in work_list]
            row_blocks, row_added, row_used_h, row_max_end = _build_one_region_row(
                sim_list, bh, region["y"], region["x0"], region["x1"]
            )
            if not row_added:
                continue
            y_after = _advance_mix_y(region["y"], row_used_h)
            if max(float(row_max_end), float(y_after)) > float(page_max_h) + 1e-9:
                continue
            max_x1 = max(float(b["x1"]) for b in row_blocks) if row_blocks else float(region["x0"])
            leftover = max(0.0, float(region["x1"]) - max_x1)
            used_area = 0.0
            for p in row_added:
                used_area += float(p["w"]) * float(p["h"])
            density = used_area / max(1.0, float(row_used_h) * float(region_w))
            score = (-density, float(leftover), float(row_used_h), -float(len(row_added)))
            if best is None or score < best_score:
                best = {
                    "base_h": float(bh),
                    "row_blocks": row_blocks,
                    "row_added": row_added,
                    "row_used_h": float(row_used_h),
                    "row_max_end": float(row_max_end),
                    "y_after": float(y_after),
                    "score": score,
                }
                best_score = score
        return best
    pages = []
    guard_pages = 0
    while True:
        types_list = [t for t in types_list if int(t["rem"]) > 0]
        if not types_list:
            break
        guard_pages += 1
        if guard_pages > 2000:
            print("⚠️ MIX_ABORT: too many pages")
            break
        page = {"blocks": [], "placements": [], "used_h": 0.0}
        regions = [
            {"name": "L", "x0": X_LEFT0, "x1": X_LEFT1, "y": 0.0},
            {"name": "R", "x0": X_RIGHT0, "x1": X_RIGHT1, "y": 0.0},
        ]
        guard_rows = 0
        while True:
            types_list = [t for t in types_list if int(t["rem"]) > 0]
            if not types_list:
                break
            guard_rows += 1
            if guard_rows > 20000:
                print("⚠️ MIX_ABORT: too many rows on one page")
                break
            row_choices = []
            for ridx, region in enumerate(regions):
                best = _find_best_region_row(types_list, region)
                if best is None:
                    continue
                other_y = float(regions[1 - ridx]["y"])
                projected_page_h = max(float(page["used_h"]), float(best["row_max_end"]), float(best["y_after"]), other_y)
                balance_gap = abs(float(best["y_after"]) - other_y)
                score = (
                    float(projected_page_h),
                    float(balance_gap),
                    float(region["y"]),
                    float(best["score"][0]),
                    float(best["score"][1]),
                    float(best["score"][2]),
                    float(best["score"][3]),
                )
                row_choices.append((score, ridx, best))
            if not row_choices:
                break
            row_choices.sort(key=lambda it: it[0])
            _row_score, ridx, best = row_choices[0]
            region = regions[ridx]
            row_blocks2, row_added2, row_used_h2, row_max_end2 = _build_one_region_row(
                types_list, best["base_h"], region["y"], region["x0"], region["x1"]
            )
            if not row_added2:
                break
            page["blocks"].extend(row_blocks2)
            page["placements"].extend(row_added2)
            region["y"] = _advance_mix_y(region["y"], row_used_h2)
            page["used_h"] = max(
                page["used_h"],
                float(row_max_end2),
                float(region["y"]),
                float(regions[1 - ridx]["y"]),
            )
            if min(float(r["y"]) for r in regions) >= float(page_max_h) - 1e-6:
                break
        if not page["placements"]:
            types_list = [t for t in types_list if int(t["rem"]) > 0]
            if not types_list:
                break
            types_list.sort(key=lambda a: (-float(a.get("W", 0.0)), -float(a.get("H", 0.0))))
            print("⚠️ MIX_SKIP_UNPLACEABLE:", types_list[0].get("tid"))
            types_list[0]["rem"] = 0
            continue
        _mix_hole_fill(page, types_list, page_max_h, gap=float(CELL_GAP_MM))
        page["used_h"] = max(page["used_h"], max(float(r["y"]) for r in regions))
        page["used_h"] = float(max(30.0, min(float(page_max_h), float(page["used_h"]))))
        pages.append(page)
    return pages
class _MaxRectsBinSimple(object):
    def __init__(self, W, H):
        self.W = int(W)
        self.H = int(H)
        self.free = [(0, 0, self.W, self.H)]
    def _prune(self, rects):
        cleaned = []
        for (x, y, w, h) in rects:
            x = int(x)
            y = int(y)
            w = int(w)
            h = int(h)
            if w <= 0 or h <= 0:
                continue
            if x < 0 or y < 0:
                continue
            if x + w > self.W or y + h > self.H:
                continue
            cleaned.append((x, y, w, h))
        out = []
        for i in range(len(cleaned)):
            xi, yi, wi, hi = cleaned[i]
            contained = False
            for j in range(len(cleaned)):
                if i == j:
                    continue
                xj, yj, wj, hj = cleaned[j]
                if xi >= xj and yi >= yj and xi + wi <= xj + wj and yi + hi <= yj + hj:
                    contained = True
                    break
            if not contained:
                out.append(cleaned[i])
        out.sort(key=lambda t: (t[1], t[0], -(t[2] * t[3])))
        return out
    def cut_out(self, ox, oy, ow, oh):
        ox = int(ox)
        oy = int(oy)
        ow = int(ow)
        oh = int(oh)
        new_free = []
        for (fx, fy, fw, fh) in self.free:
            if (ox >= fx + fw) or (ox + ow <= fx) or (oy >= fy + fh) or (oy + oh <= fy):
                new_free.append((fx, fy, fw, fh))
                continue
            if oy > fy:
                new_free.append((fx, fy, fw, oy - fy))
            if oy + oh < fy + fh:
                new_free.append((fx, oy + oh, fw, (fy + fh) - (oy + oh)))
            top = max(fy, oy)
            bot = min(fy + fh, oy + oh)
            hh = bot - top
            if hh > 0 and ox > fx:
                new_free.append((fx, top, ox - fx, hh))
            if hh > 0 and ox + ow < fx + fw:
                new_free.append((ox + ow, top, (fx + fw) - (ox + ow), hh))
        self.free = self._prune(new_free)
    def find_bottom_left(self, rw, rh):
        rw = int(rw)
        rh = int(rh)
        best = None
        for (fx, fy, fw, fh) in self.free:
            if rw <= fw and rh <= fh:
                short_side = min(fw - rw, fh - rh)
                long_side = max(fw - rw, fh - rh)
                waste = (fw * fh) - (rw * rh)
                cand = (short_side, long_side, waste, fy, fx, fw, fh)
                if best is None or cand < best:
                    best = cand
        if best is None:
            return None
        return {"x": best[4], "y": best[3]}
def _mix_hole_fill(page, types_list, page_max_h, gap):
    H = float(page.get("used_h", 0.0))
    if H <= 1e-6:
        return
    usable_x0 = float(PAGE_MARGIN_L)
    usable_x1 = float(PAGE_W - PAGE_MARGIN_R)
    usable_y0 = 0.0
    usable_y1 = float(page_max_h)
    binW = int(round(usable_x1 - usable_x0))
    binH = int(round(usable_y1 - usable_y0))
    bp = _MaxRectsBinSimple(binW, binH)
    block_x = int(round(SPLIT_X - usable_x0))
    block_w = int(round(SPLIT_GAP_W))
    bp.cut_out(block_x, 0, block_w, binH)
    for b in page.get("blocks", []):
        x = float(b["x0"]) - usable_x0
        y = float(b["y0"]) - usable_y0
        w = float(b["x1"]) - float(b["x0"])
        h = float(b["y1"]) - float(b["y0"])
        ox = int(round(x))
        oy = int(round(y))
        ow = int(round(w + gap))
        oh = int(round(h + gap))
        bp.cut_out(ox, oy, ow, oh)
    for p in page.get("placements", []):
        x = float(p["x"]) - usable_x0
        y = float(p["y"]) - usable_y0
        w = float(p["w"])
        h = float(p["h"])
        ox = int(round(x))
        oy = int(round(y))
        ow = int(round(w + gap))
        oh = int(round(h + gap))
        bp.cut_out(ox, oy, ow, oh)
    cand_types = [t for t in types_list if int(t.get("rem", 0)) > 0]
    cand_types.sort(key=lambda t: -(float(t["W"]) * (float(t["H"]) + float(QR_BAND))))
    added = 0
    guard = 0
    while True:
        guard += 1
        if guard > 200000:
            break
        placed_any = False
        for t in cand_types:
            if int(t["rem"]) <= 0:
                continue
            w0 = float(t["W"])
            h0 = float(t["H"])
            opts = _iter_mix_orients(w0, h0, bool(t.get("rot90", False)))
            def _lane_rank(w):
                if w <= RIGHT_BLOCK_W + 1e-9:
                    return 0
                if w <= LEFT_BLOCK_W + 1e-9:
                    return 1
                return 2
            opts.sort(key=lambda x: (_lane_rank(x[0]), x[0], x[1]))
            best_pick = None
            for (wi, hi, rot90) in opts:
                total_h = float(QR_BAND) + float(hi)
                rw = int(round(wi))
                rh = int(round(total_h))
                pos = bp.find_bottom_left(rw + int(round(gap)), rh + int(round(gap)))
                if pos is None:
                    continue
                rank = (
                    int(pos["y"]),
                    int(pos["x"]),
                    -int(round(wi * total_h)),
                    abs(int(round(wi)) - int(round(total_h))),
                )
                if best_pick is None or rank < best_pick[0]:
                    best_pick = (rank, pos, wi, hi, total_h, rot90)
            if best_pick is None:
                continue
            _rank, pos, wi, hi, total_h, rot90 = best_pick
            bp.cut_out(pos["x"], pos["y"], int(round(wi + gap)), int(round(total_h + gap)))
            block_x0 = float(pos["x"]) + usable_x0
            block_y0 = float(pos["y"]) + usable_y0
            img_y = float(block_y0) + float(QR_BAND)
            page.setdefault("blocks", []).append({
                "type": t,
                "tid": t["tid"],
                "x0": float(block_x0),
                "x1": float(block_x0) + float(wi),
                "y0": float(block_y0),
                "y1": float(img_y) + float(hi),
            })
            page["placements"].append({
                "type": t,
                "tid": t["tid"],
                "x": float(block_x0),
                "y": float(img_y),
                "w": float(wi),
                "h": float(hi),
                "rot90": bool(rot90),
            })
            t["rem"] -= 1
            added += 1
            placed_any = True
            break
        if not placed_any:
            break
    if added > 0:
        max_end = 0.0
        for b in page.get("blocks", []):
            max_end = max(max_end, float(b["y1"]))
        for p in page["placements"]:
            max_end = max(max_end, float(p["y"]) + float(p["h"]))
        page["used_h"] = max(page.get("used_h", 0.0), max_end)
def _mix_draw_edge_ticks(page, items, page_h_mm):
    if not items:
        return
    eps = 1.0
    blocks = []
    for it in items:
        if ("x0" in it) and ("x1" in it) and ("y0" in it) and ("y1" in it):
            blocks.append({
                "x0": float(it["x0"]),
                "x1": float(it["x1"]),
                "y0": float(it["y0"]),
                "y1": float(it["y1"]),
            })
        elif ("x" in it) and ("y" in it) and ("w" in it) and ("h" in it):
            x0 = float(it["x"])
            y0 = float(it["y"])
            x1 = x0 + float(it["w"])
            y1 = y0 + float(it["h"])
            blocks.append({
                "x0": x0,
                "x1": x1,
                "y0": y0,
                "y1": y1,
            })
    if not blocks:
        return
    min_y0 = min(float(b["y0"]) for b in blocks)
    top_blocks = []
    for b in blocks:
        if abs(float(b["y0"]) - min_y0) <= eps:
            top_blocks.append(b)
    xs = set()
    for b in top_blocks:
        xs.add(round(float(b["x0"]), 3))
        xs.add(round(float(b["x1"]), 3))
    for x in sorted(xs):
        if 0.0 - 1e-6 <= x <= float(PAGE_W) + 1e-6:
            _draw_top_edge_tick(page, x, tick_len_mm=MARK_LEN, lw=1.0)
    min_x0 = min(float(b["x0"]) for b in blocks)
    left_blocks = []
    for b in blocks:
        if abs(float(b["x0"]) - min_x0) <= eps:
            left_blocks.append(b)
    ys = set()
    for b in left_blocks:
        top_cut_y = float(b["y0"]) + float(QR_BAND) + 0.8
        bottom_cut_y = float(b["y1"])
        if 0.0 - 1e-6 <= top_cut_y <= float(page_h_mm) + 1e-6:
            ys.add(round(top_cut_y, 3))
        if 0.0 - 1e-6 <= bottom_cut_y <= float(page_h_mm) + 1e-6:
            ys.add(round(bottom_cut_y, 3))
    for y in sorted(ys):
        if 0.0 - 1e-6 <= y <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, y, tick_len_mm=MARK_LEN, lw=1.0)
def render_mix_page(page_obj, is_contour):
    H = float(page_obj.get("used_h", 30.0)) + float(MARGIN_BOTTOM_MM) + float(SAFE_PAD_MM)
    W = float(PAGE_W)
    doc = fitz.open()
    page = doc.new_page(width=mm_to_pt(W), height=mm_to_pt(H))
    page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=(0, 0, 0), width=1.0)
    for b in page_obj.get("blocks", []):
        t = b["type"]
        x0 = float(b["x0"])
        x1 = float(b["x1"])
        y0 = float(b["y0"])
        qr_bytes = t.get("qr_bytes", None)
        if qr_bytes:
            qr_rect = fitz.Rect(
                mm_to_pt(x1 - QR_W), mm_to_pt(y0),
                mm_to_pt(x1), mm_to_pt(y0 + QR_H)
            )
            page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)
        label_text = t.get("label_text", "")
        if label_text:
            _draw_label_top_left(page, label_text, x0, y0, max_w_mm=max(10.0, (x1 - x0 - QR_W)))
    placements = page_obj.get("placements", [])
    for p in placements:
        t = p["type"]
        img_bytes = t["img_cont"] if is_contour else t["img_body"]
        x = float(p["x"])
        y = float(p["y"])
        w = float(p["w"])
        h = float(p["h"])
        rot90 = bool(p.get("rot90", False))
        outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
        if DRAW_PART_OUTER_BOX:
            page.draw_rect(outer, color=(0, 0, 0), width=0.5)
        if img_bytes:
            page.insert_image(outer, stream=img_bytes, keep_proportion=True, rotate=(90 if rot90 else 0))
    _mix_draw_edge_ticks(page, page_obj.get("blocks", []), page_h_mm=H)
    return doc
def _run_mix_stage(mix_types, progress_cb=None, log_cb=None, ok_list=None, skip_list=None, progress_total=None):
    mix_outputs_p1 = []
    mix_outputs_p2 = []
    ok_list = list(ok_list or [])
    skip_list = list(skip_list or [])
    n_total = int(progress_total) if progress_total is not None else max(1, len(mix_types))
    if progress_cb:
        progress_cb(0, n_total, "混拼排版中")
    if mix_types:
        mix_pages = pack_mix_by_height_rule(mix_types, page_max_h=PAGE_H_MAX)
        mix_pages = [pg for pg in mix_pages if (pg.get("placements") or [])]
        for i, pg in enumerate(mix_pages, start=1):
            name = "mix_%d.pdf" % i
            out_path1 = os.path.join(DEST_DIR1, name)
            out_path2 = os.path.join(DEST_DIR2, name)
            doc1 = render_mix_page(pg, is_contour=False)
            doc2 = render_mix_page(pg, is_contour=True)
            p1, p2 = _save_dual_docs(doc1, doc2, out_path1, out_path2)
            mix_outputs_p1.append(p1)
            mix_outputs_p2.append(p2)
            _log(log_cb, "🧩 %s blocks=%d items=%d used_h=%.1fmm  %s | %s" %
                 (name, len(pg.get("blocks") or []), len(pg.get("placements") or []), float(pg.get("used_h", 0.0)), p1, p2))
    _finish_stage_progress(progress_cb, n_total)
    return {
        "mix_p1_files": mix_outputs_p1,
        "mix_p2_files": mix_outputs_p2,
        "ok": ok_list,
        "skip": skip_list,
    }
def _wrap_stage_progress(progress_cb, stage_idx):
    if progress_cb is None:
        return None
    def _cb(cur, tot, msg):
        try:
            cur_i = int(cur)
            tot_i = max(1, int(tot))
        except Exception:
            cur_i = 0
            tot_i = 1
        cur_i = max(0, min(cur_i, tot_i))
        progress_cb(stage_idx * tot_i + cur_i, 2 * tot_i, msg)
    return _cb
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
    archived_paths = _collect_archived_paths(input_pdfs=input_pdfs, log_cb=log_cb)
    _log(log_cb, "=== STAGE1/2 整拼（single）===")
    res_single = _run_single_stage(
        archived_paths,
        progress_cb=_wrap_stage_progress(progress_cb, 0),
        log_cb=log_cb
    )
    _log(log_cb, "=== STAGE2/2 混拼（mix）===")
    mix_pool = res_single.get("mix_pool") or []
    res_mix = _run_mix_stage(
        mix_pool,
        progress_cb=_wrap_stage_progress(progress_cb, 1),
        log_cb=log_cb,
        ok_list=res_single.get("ok_mix_pool") or [],
        skip_list=res_single.get("skip_mix_pool") or [],
        progress_total=len(archived_paths),
    )
    if progress_cb:
        total = max(1, len(archived_paths))
        progress_cb(2 * total, 2 * total, "完成 %d/%d(整拼+混拼)" % (2 * total, 2 * total))
    return {
        "single_p1_files": res_single.get("single_p1_files") or [],
        "single_p2_files": res_single.get("single_p2_files") or [],
        "mix_p1_files": res_mix.get("mix_p1_files") or [],
        "mix_p2_files": res_mix.get("mix_p2_files") or [],
        "ok_single": res_single.get("ok") or [],
        "skip_single": res_single.get("skip") or [],
        "ok_mix": res_mix.get("ok") or [],
        "skip_mix": res_mix.get("skip") or [],
        "ok": (res_single.get("ok") or []) + (res_mix.get("ok") or []),
        "skip": (res_single.get("skip") or []) + (res_mix.get("skip") or []),
    }
def main():
    res = run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None)
    print("DONE:")
    print("  single_p1_files:", len(res.get("single_p1_files") or []))
    print("  single_p2_files:", len(res.get("single_p2_files") or []))
    print("  mix_p1_files:", len(res.get("mix_p1_files") or []))
    print("  mix_p2_files:", len(res.get("mix_p2_files") or []))
if __name__ == "__main__":
    main()
                                                                                                       