# -*- coding: utf-8 -*-
import fitz
import numpy as np
from PIL import Image

import gb5_config as C
from gb5_utils import clamp_bbox, mm_to_px


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

    from gb5_utils import rect_union
    return {
        "page_rect": page_rect,
        "text_boxes": text_boxes,
        "image_bbox": rect_union(image_boxes),
        "draw_bbox": rect_union(draw_rects),
    }


def render_page_to_pil(pdf_path, page_index=0, dpi=C.RENDER_DPI, doc=None):
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


def mask_text_regions_on_pil(pil_img, text_boxes_pt, page_rect_pt, pad_mm=C.TEXT_MASK_PAD_MM):
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