# -*- coding: utf-8 -*-
import numpy as np
from io import BytesIO
from PIL import Image

import gb5_config as C
from gb5_utils import clamp_bbox, union_bbox, mm_to_px
from gb5_pdf_struct import (
    get_page_struct_info, render_page_to_pil, pdf_rect_to_px_bbox, mask_text_regions_on_pil
)

if C.CV2_OK:
    import cv2


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

    if not C.CV2_OK:
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
        th = cv2.adaptiveThreshold(
            blur, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            35, 5
        )
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

    try:
        for thr in [250, 245, 240]:
            mask = (gray < thr).astype(np.uint8) * 255
            mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8), iterations=1)
            b = _bbox_from_mask(mask, W, H)
            if b:
                candidates.append(b)
                break
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

    masked = mask_text_regions_on_pil(pil_img, text_boxes, page_rect, pad_mm=C.TEXT_MASK_PAD_MM)
    cv_bbox_px = find_outer_bbox(masked)

    return draw_bbox_px, image_bbox_px, cv_bbox_px


def _content_touches_edges(pil_rgba, thr=250, margin_px=2):
    arr = np.array(pil_rgba.convert("RGB"))
    gray = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]).astype(np.uint8)
    m = max(1, int(margin_px))
    mask = (gray < thr)
    if mask[:m, :].any():
        return True
    if mask[-m:, :].any():
        return True
    if mask[:, :m].any():
        return True
    if mask[:, -m:].any():
        return True
    return False


def _expand_bbox_clamped(bbox, exp_px, img_w, img_h):
    x0, y0, x1, y1 = bbox
    return clamp_bbox(x0 - exp_px, y0 - exp_px, x1 + exp_px, y1 + exp_px, img_w, img_h)


def make_part_png_bytes_using_ref_bbox(pdf_path, page_index, ref_bbox, ref_size, dpi=C.RENDER_DPI, doc=None):
    pad_px = mm_to_px(C.CROP_PAD_MM, dpi)

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
            bbox_ref_scaled = (int(round(x0 * sx)), int(round(y0 * sy)), int(round(x1 * sx)), int(round(y1 * sy)))

    bbox = union_bbox(bbox_ref_scaled, bbox_self)
    if bbox is None:
        bbox = (0, 0, img_w, img_h)

    x0, y0, x1, y1 = bbox
    x0, y0, x1, y1 = clamp_bbox(x0 - pad_px, y0 - pad_px, x1 + pad_px, y1 + pad_px, img_w, img_h)

    for _ in range(2):
        crop = img.crop((x0, y0, x1, y1)).convert("RGBA")
        if not _content_touches_edges(crop, thr=250, margin_px=2):
            break
        extra_px = mm_to_px(0.6, dpi)
        x0, y0, x1, y1 = _expand_bbox_clamped((x0, y0, x1, y1), extra_px, img_w, img_h)

    bio = BytesIO()
    img.crop((x0, y0, x1, y1)).convert("RGBA").save(bio, format="PNG", optimize=True)
    return bio.getvalue()