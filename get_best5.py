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


# =========================
# 路径（可由 run.py 动态注入）
# =========================
DEST_DIR = r"D:\WXWork\1688857155804169\Cache\File\2026-02\test1"       # 原始输入 PDF 所在目录
IN_PDF_ARCHIVE_DIR = r"D:\WXWork\1688857155804169\Cache\File\2026-02\test1"  # 归档后的 PDF 目录（先复制再处理）
DEST_DIR1 = r"D:\WXWork\1688857155804169\Cache\File\2026-02\test1"      # 输出图形页 PDF 的目录
DEST_DIR2 = r"D:\WXWork\1688857155804169\Cache\File\2026-02\test1"      # 输出轮廓页 PDF 的目录

def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):       
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)
OUTER_EXT = 2.0               # 单个图块外扩边距(mm)，即尺寸 A/B 两边各加 2mm
CELL_GAP_MM = 10.0            # 图与图之间的间距(mm)
MARGIN_LR_MM = 5.0            # 纸张左右边距(mm)
MARGIN_TOP_MM = 10.0          # 纸张顶部边距(mm)
MARGIN_BOTTOM_MM = 5.0        # 纸张底部边距(mm)
INNER_MARGIN_MM = 0.0         # 图像插入到外框时的内缩边距(mm)，当前不内缩
MARK_LEN = 5.0                # 刀线长度(mm)
PAGE_W = 600.0                # 整张纸宽度(mm)
PAGE_H_MAX = 1500.0           # 单页最大高度(mm)
PAGE_MARGIN_L = MARGIN_LR_MM  # 实际用于排版计算的左边距(mm)
PAGE_MARGIN_R = MARGIN_LR_MM  # 实际用于排版计算的右边距(mm)
SPLIT_X = 320.0               # 中间分区线 x=320(mm)
SPLIT_GAP_W = 10.0            # 中缝空带宽度(mm)，即 320~330 这段不放图
QR_BAND = MARGIN_TOP_MM       # 每个块顶部预留带高度(mm)，用于放二维码和标签
QR_W = 10.0                   # 二维码宽度(mm)
QR_H = 10.0                   # 二维码高度(mm)
LABEL_FONT_SIZE = 10          # 标签字号(pt)
RENDER_DPI = 600              # PDF 转图片时的渲染分辨率
CROP_PAD_MM = 1.2             # 裁剪 bbox 时额外外扩的安全边距(mm)
TEXT_MASK_PAD_MM = 1.2        # 抹掉文字区域时的额外扩张边距(mm)
SAFE_PAD_MM = 3.0             # 页高裁切时的额外保险高度(mm)，防止底部缺图
DRAW_PART_OUTER_BOX = False   # 是否在每个小图外再画一个外框（调试用）
MIX_ENABLE = True             # 是否启用混拼
MIX_THRESHOLD = 10            # 数量阈值：N < 10 才混拼；N >= 10 直接跳过
LEFT_BLOCK_W = float(SPLIT_X - PAGE_MARGIN_L)               # 左侧可用最大不跨区块宽：320-5=315(mm)
RIGHT_BLOCK_W = float((PAGE_W - PAGE_MARGIN_R) - (SPLIT_X + SPLIT_GAP_W))  # 右侧可用最大不跨区块宽：595-330=265(mm)
TOTAL_USABLE_W = float((PAGE_W - PAGE_MARGIN_R) - PAGE_MARGIN_L)  # 整张纸真正可用宽度：600-5-5=590(mm)
EDGE_SNAP_MM = 10.0           # 吸边容差(mm)，当前版本基本未启用，预留参数
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

def _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0):
    ypt = mm_to_pt(float(y_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(0, ypt), fitz.Point(tick, ypt), color=(0, 0, 0), width=lw)

def _draw_top_edge_tick(page, x_mm, tick_len_mm=MARK_LEN, lw=1.0):
    xpt = mm_to_pt(float(x_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, tick), color=(0, 0, 0), width=lw)

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
            version=None,                                      # 自动选择二维码版本
            error_correction=qrcode.constants.ERROR_CORRECT_M, # 中等级别容错
            box_size=8,                                        # 二维码单模块像素大小
            border=0                                           # 不额外留白边框
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
    est = 0.55 * fontsize * len(text)  # 粗略估算文字宽度
    if est <= max_width_pt:
        return text
    keep = max(0, int(max_width_pt / (0.55 * fontsize)) - 3)
    if keep <= 0:
        return "..."
    return text[:keep] + "..."

def _draw_label_top_left(page, text, x_left_mm, y_top_mm, max_w_mm):
    fontfile = pick_cjk_fontfile()
    fontsize = LABEL_FONT_SIZE            # 标签字号
    max_w_pt = mm_to_pt(max(5.0, float(max_w_mm)))  # 标签最大可绘制宽度
    txt = trim_text_to_width(text, fontsize, max_w_pt)
    xpt = mm_to_pt(float(x_left_mm) + 1.0)  # 相对块左边再右移 1mm
    ypt = mm_to_pt(float(y_top_mm) + 7.5)   # 相对块顶再下移 7.5mm
    if fontfile:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontname="helv", color=(0, 0, 0))

def _is_probably_page_frame(rect_obj, page_rect):
    if rect_obj is None:
        return False
    tol = 2.0  # 允许的边界误差(pt)
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
                text_boxes.append(r)  # 文本框
            elif btype == 1:
                if not _is_probably_page_frame(r, page_rect):
                    image_boxes.append(r)  # 图像块
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
            draw_rects.append(r)  # 矢量绘图区域
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
    sx = img_w / float(page_rect_pt.width)   # x 方向 pt->px 比例
    sy = img_h / float(page_rect_pt.height)  # y 方向 pt->px 比例

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

    pad_px_x = max(1, int(round((pad_mm * 72.0 / 25.4) * sx)))  # 文本框 x 方向额外扩张像素
    pad_px_y = max(1, int(round((pad_mm * 72.0 / 25.4) * sy)))  # 文本框 y 方向额外扩张像素

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
    edge_pad = 8  # 贴边容差像素
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
        if float(area) >= 0.08 * max_area:  # 只保留面积足够大的轮廓
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
        for thr in [250, 245, 240, 235]:  # 多个阈值尝试
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
    pad_px = mm_to_px(CROP_PAD_MM, dpi)  # 初始裁剪安全边距像素

    img = render_page_to_pil(pdf_path, page_index=page_index, dpi=dpi, doc=doc)
    img_w, img_h = img.size

    draw_px, image_px, cv_px = get_page_bbox_candidates_px(doc, page_index, img)
    bbox_self = union_bbox(union_bbox(draw_px, image_px), cv_px)

    bbox_ref_scaled = None
    if ref_bbox is not None and ref_size is not None:
        rw, rh = ref_size  # 参考页图像尺寸
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

    # 额外检查是否裁掉了边缘内容，必要时再扩一次
    for _ in range(2):
        crop = img.crop((x0, y0, x1, y1)).convert("RGBA")
        arr = np.array(crop.convert("RGB"))
        gray = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]).astype(np.uint8)
        m = 2  # 检查边缘 2 像素内是否还有内容
        if (gray[:m, :] < 250).any() or (gray[-m:, :] < 250).any() or (gray[:, :m] < 250).any() or (gray[:, -m:] < 250).any():
            extra_px = mm_to_px(0.6, dpi)
            x0, y0, x1, y1 = _expand_bbox_clamped((x0, y0, x1, y1), extra_px, img_w, img_h)
        else:
            break

    bio = BytesIO()
    img.crop((x0, y0, x1, y1)).convert("RGBA").save(bio, format="PNG", optimize=True)
    return bio.getvalue()

def _mix_choose_orient(w0, h0):
    w0 = float(w0)
    h0 = float(h0)

    w_a, h_a, r_a = w0, h0, False   # 原始方向
    w_b, h_b, r_b = h0, w0, True    # 旋转90度方向

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
        y2 += 10.0  # 跨过 464 高度边界时插入 10mm 缝

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
    X_LEFT0 = float(PAGE_MARGIN_L)             # 左区起排 x=5
    X_LEFT1 = float(SPLIT_X)                   # 左区终点 x=320
    X_RIGHT0 = float(SPLIT_X + SPLIT_GAP_W)    # 右区起点 x=330
    X_RIGHT1 = float(PAGE_W - PAGE_MARGIN_R)   # 右区终点 x=595
    gap = float(CELL_GAP_MM)                   # 小图之间的间距

    types_list = []
    for t in mix_types:
        tt = dict(t)
        tt["rem"] = int(tt.get("rem", tt.get("N", 0)) or 0)  # rem=当前还剩多少张未排
        types_list.append(tt)

    # 若某个块在左右任一栏都放不下，则直接视为不可排
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
            rem_w = float(x_end) - float(x)  # 当前栏剩余可用宽度
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

            x0 = float(x)  # 这个大块的左边界
            block_w = float(cand_plan["block_w"])
            x1 = x0 + block_w  # 这个大块的右边界

            block_min_x = None
            block_max_x = None
            block_max_y = None

            for c in range(cols):
                col_x = x0 + float(c) * (Wi + gap)  # 当前列左上角 x
                for s in range(stack_cap):
                    if int(cand["rem"]) <= 0:
                        break
                    px = col_x
                    py = float(y) + float(QR_BAND) + float(s) * (Hi + gap)  # 每张小图的实际 y，从顶部带下面开始排

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
                    "y0": float(y),           # 大块顶部，标签/二维码贴这里
                    "y1": float(block_max_y), # 大块底部
                })

            x = x1

        if not row_added:
            return [], [], 0.0, 0.0

        row_max_end = 0.0
        for p in row_added:
            row_max_end = max(row_max_end, float(p["y"]) + float(p["h"]))
        row_used_h = max(float(QR_BAND), row_max_end - float(y))  # 这一行总高度=顶部带+图像区
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
        self.free = [(0, 0, self.W, self.H)]  # 当前空闲矩形列表

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
        rw = int(rw)  # 待放矩形宽
        rh = int(rh)  # 待放矩形高
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

    usable_x0 = float(PAGE_MARGIN_L)            # 真正可放区域左边界
    usable_x1 = float(PAGE_W - PAGE_MARGIN_R)   # 真正可放区域右边界
    usable_y0 = 0.0
    usable_y1 = float(page_max_h)

    binW = int(round(usable_x1 - usable_x0))    # 空洞填充器宽度
    binH = int(round(usable_y1 - usable_y0))    # 空洞填充器高度
    bp = _MaxRectsBinSimple(binW, binH)

    block_x = int(round(SPLIT_X - usable_x0))   # 中缝障碍物左边界
    block_w = int(round(SPLIT_GAP_W))           # 中缝障碍物宽度
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
        ow = int(round(w + gap))   # 把 gap 也算进占用，避免补洞后贴太近
        oh = int(round(h + gap))
        bp.cut_out(ox, oy, ow, oh)

    cand_types = [t for t in types_list if int(t.get("rem", 0)) > 0]
    cand_types.sort(key=lambda t: -(float(t["W"]) * (float(t["H"]) + float(QR_BAND))))  # 优先补大块

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

    eps = 1.0  # 判断“是否同一批顶部/左侧块”的容差(mm)

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

    # 顶部块：画顶部刀线
    min_y0 = min(float(b["y0"]) for b in blocks)  # 整页最顶部块的 y0
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

    # 左侧块：画左侧刀线
    min_x0 = min(float(b["x0"]) for b in blocks)  # 整页最左块的 x0
    left_blocks = []
    for b in blocks:
        if abs(float(b["x0"]) - min_x0) <= eps:
            left_blocks.append(b)

    ys = set()
    for b in left_blocks:
        top_cut_y = float(b["y0"]) + float(QR_BAND) + 0.8  # 左侧刀线上端：放到图片区顶部稍下
        bottom_cut_y = float(b["y1"])                      # 左侧刀线下端：块底部

        if 0.0 - 1e-6 <= top_cut_y <= float(page_h_mm) + 1e-6:
            ys.add(round(top_cut_y, 3))
        if 0.0 - 1e-6 <= bottom_cut_y <= float(page_h_mm) + 1e-6:
            ys.add(round(bottom_cut_y, 3))
    for y in sorted(ys):
        if 0.0 - 1e-6 <= y <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, y, tick_len_mm=MARK_LEN, lw=1.0)

def render_mix_page(page_obj, is_contour):
    H = float(page_obj.get("used_h", 30.0)) + float(MARGIN_BOTTOM_MM) + float(SAFE_PAD_MM)  # 实际页高
    W = float(PAGE_W)  # 页宽固定 600mm

    doc = fitz.open()
    page = doc.new_page(width=mm_to_pt(W), height=mm_to_pt(H))
    page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=(0, 0, 0), width=1.0)

    # 先画二维码和标签（位于块顶部预留带）
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

    # 再贴实际图片
    placements = page_obj.get("placements", [])
    for p in placements:
        t = p["type"]
        img_bytes = t["img_cont"] if is_contour else t["img_body"]  # 图形页/轮廓页切换

        x = float(p["x"])
        y = float(p["y"])
        w = float(p["w"])
        h = float(p["h"])
        rot90 = bool(p.get("rot90", False))

        outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
        if DRAW_PART_OUTER_BOX:
            page.draw_rect(outer, color=(0, 0, 0), width=0.5)

        inner = outer
        if img_bytes:
            page.insert_image(inner, stream=img_bytes, keep_proportion=True, rotate=(90 if rot90 else 0))
    _mix_draw_edge_ticks(page, page_obj.get("blocks", []), page_h_mm=H)
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

    mix_types = []
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
            outer_w0 = float(A + 2.0 * OUTER_EXT)  # 图块外扩后宽度
            outer_h0 = float(B + 2.0 * OUTER_EXT)  # 图块外扩后高度

            if int(N) >= MIX_THRESHOLD:
                skip_list.append((type_name_raw, "skip_N_ge_%d" % MIX_THRESHOLD))
                _log(log_cb, "⏭️ SKIP(N>=%d): %s" % (MIX_THRESHOLD, type_name_raw))
                continue

            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)
            qr_bytes = make_qr_png_bytes(qr_text)

            for pi in range(pair_count):
                page_body = 2 * pi      # 图形页页号
                page_cont = 2 * pi + 1  # 轮廓页页号

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

                tid = "%s@P%d" % (type_name_raw, pi + 1)  # 每个页对单独视为一个 type id

                Mi, Ni, rot90 = _mix_choose_orient(outer_w0, outer_h0)
                mix_types.append({
                    "tid": tid,                # 类型唯一 ID
                    "label_text": label_text,  # 标签文字
                    "qr_bytes": qr_bytes,      # 二维码 PNG 数据
                    "img_body": img_body,      # 图形页 PNG
                    "img_cont": img_cont,      # 轮廓页 PNG
                    "rem": int(N),             # 还需摆放的张数
                    "W": float(Mi),            # 当前朝向宽度
                    "H": float(Ni),            # 当前朝向高度
                    "rot90": bool(rot90),      # 是否旋转90度
                })
                ok_list.append((tid, "mix_pool"))
                _log(log_cb, "🧩 MIX_POOL: %s N=%d Mi=%.1f Ni=%.1f rot90=%s" %
                     (tid, int(N), float(Mi), float(Ni), str(bool(rot90))))

        except Exception as e:
            skip_list.append((type_name_raw, repr(e)))
            _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, repr(e)))
        finally:
            try:
                if d is not None:
                    d.close()
            except Exception:
                pass

    mix_outputs_p1 = []
    mix_outputs_p2 = []
    if mix_types:
        mix_pages = pack_mix_by_height_rule(mix_types, page_max_h=PAGE_H_MAX)
        mix_pages = [pg for pg in mix_pages if (pg.get("placements") or [])]

        for i, pg in enumerate(mix_pages, start=1):
            name = "mix_%d.pdf" % i
            out_path1 = os.path.join(DEST_DIR1, name)
            out_path2 = os.path.join(DEST_DIR2, name)

            doc1 = render_mix_page(pg, is_contour=False)  # 图形页
            doc2 = render_mix_page(pg, is_contour=True)   # 轮廓页

            p1 = safe_save(doc1, out_path1)
            p2 = safe_save(doc2, out_path2)
            mix_outputs_p1.append(p1)
            mix_outputs_p2.append(p2)

            _log(log_cb, "🧩 %s blocks=%d items=%d used_h=%.1fmm  %s | %s" %
                 (name, len(pg.get("blocks") or []), len(pg.get("placements") or []), float(pg.get("used_h", 0.0)), p1, p2))

    if progress_cb:
        progress_cb(N_total, N_total, "完成 %d/%d" % (N_total, N_total))

    return {
        "mix_p1_files": mix_outputs_p1,
        "mix_p2_files": mix_outputs_p2,
        "ok": ok_list,
        "skip": skip_list,
    }

def main():
    res = run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None)
    print("DONE:")
    print("  mix_p1_files:", len(res.get("mix_p1_files") or []))
    print("  mix_p2_files:", len(res.get("mix_p2_files") or []))
    print("  skip:", len(res.get("skip") or []))

if __name__ == "__main__":
    main()