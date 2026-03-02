# -*- coding: utf-8 -*-
import os
import time
import shutil
import hashlib

import fitz
import gb5_config as C


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
    return (min(b1[0], b2[0]), min(b1[1], b2[1]), max(b1[2], b2[2]), max(b1[3], b2[3]))


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
    if len(s) <= max_len:
        return s
    h = hashlib.md5(s.encode("utf-8", "ignore")).hexdigest()[:10]
    keep = max(40, max_len - 11)
    return s[:keep] + "_" + h


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


def safe_save(doc, out_path):
    ensure_dir(os.path.dirname(out_path))
    tmp_pdf = os.path.join(os.path.dirname(out_path), os.path.splitext(os.path.basename(out_path))[0] + "_tmp.pdf")
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
        alt_pdf = os.path.join(os.path.dirname(out_path), os.path.splitext(os.path.basename(out_path))[0] + "_%s.pdf" % ts)
        try:
            shutil.move(tmp_pdf, alt_pdf)
        except Exception:
            shutil.copyfile(tmp_pdf, alt_pdf)
            try:
                os.remove(tmp_pdf)
            except Exception:
                pass
        return alt_pdf