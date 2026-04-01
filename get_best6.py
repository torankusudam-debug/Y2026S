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
DEST_DIR = r"D:\我的数据\文档\get"
IN_PDF_ARCHIVE_DIR = r"D:\我的数据\文档\1"
DEST_DIR1 = r"D:\我的数据\文档\2"
DEST_DIR2 = r"D:\我的数据\文档\3"
OUTER_EXT = 2.0                             # 刀线外扩（mm），避免裁剪时漏掉边缘内容
INNER_MARGIN_MM = 3.0                       # 内边距（mm），避免刀线太靠近内容，影响美观和实用, 也避免裁剪时误伤内容。
MARK_LEN = 5.0                              # 刀线/刻度长度（mm）
SINGLE_CUT_OUTER_EXT_MM = 2.5               # 单切时原图四边外扩；相邻原图间距 = 2 * 该值
SINGLE_CUT_SHARED_LINE_TRIM_MM = 0.35       # 单切刀版共线去重时，重复边裁掉的宽度(mm)；过大容易缺边，过小会看起来像双线。
SINGLE_CUT_CONTOUR_LINE_W_PT = 0.55         # 单切刀版绿色轮廓线的线宽(pt)。
CONTOUR_LINE_EXT_MM = 2.0                   # 刀版格线超出端头的长度(mm)。
SINGLE_SINGLE_LINE_SPECIAL_MODE = True      # 是否启用“单枚单线”单拼特殊规则；False 时回到普通单拼逻辑。
SINGLE_SPECIAL_QR_SHIFT_LEFT_MM = 5.0       # 单枚单线模式下，二维码整体向左移动量(mm)。
SINGLE_SPECIAL_GAP_MARK_MIN_MM = 8.0        # 单枚单线模式下，分块缝中间保留顶部刀线时的最小安全距离(mm)。
SINGLE_SPECIAL_LABEL_X_MM = 8.0             # 单枚单线模式下，标签左上角的 X 位置(mm)，相对单拼内容区左边计算。
SINGLE_SPECIAL_LABEL_Y_MM = 2.0             # 单枚单线模式下，标签左上角的 Y 位置(mm)，相对单拼内容区顶部计算。
SINGLE_SPECIAL_BOTTOM_CLEAR_MM = 3.0        # 单枚单线模式下，拼图底部额外留白(mm)，避免内容贴住 PDF 底边。
SINGLE_LEFT_CUT_TICK_SHIFT_LEFT_MM = 5.0    # 单拼左侧刀线在旋转前向左平移量(mm)；只影响左侧那组短刀线。
SINGLE_QR_RIGHT_CUTLINE_GAP_MM = 1.0        # 二维码与最右侧顶部刀线的最小避让距离(mm)。
SINGLE_H_MAX = 800.0                        # 单拼最终 PDF 总高度上限(mm)。
SINGLE_PRINT_PAGE_W_MM = 585.0              # 单拼最终 PDF 固定宽度(mm)。
SINGLE_PRINT_SIDE_MARGIN_MM = 7.0           # 单拼最终 PDF 左右留边(mm)；矩形1 到 PDF 左右边的最小距离。
SINGLE_PRINT_TOP_MARGIN_MM = 15.0           # 单拼最终 PDF 顶部留边(mm)。
SINGLE_PRINT_BOTTOM_MARGIN_MM = 15.0        # 单拼最终 PDF 底部留边(mm)。
SINGLE_PRINT_CORNER_LEN_MM = 8.0            # 单拼印刷版四角黑色粗角线的边长(mm)。
SINGLE_PRINT_CORNER_W_PT = 2.2              # 单拼印刷版四角黑色粗角线的线宽(pt)。
SINGLE_PRINT_SEG_MARK_STEP_MM = 300.0       # 单拼印刷版侧边分段标记的纵向步长(mm)；每隔这么高画一组。
SINGLE_PRINT_SEG_MARK_EDGE_LEN_MM = 14.0    # 单拼印刷版侧边分段标记中，贴边竖线的长度(mm)。
SINGLE_PRINT_SEG_MARK_W_PT = 1.2            # 单拼印刷版侧边分段标记线宽(pt)。
SINGLE_PRINT_SEG_MARK_FONT_PT = 8.0         # 单拼印刷版左侧编号字体大小(pt)。
SINGLE_PRINT_SEG_NUMBER_SHIFT_DOWN_MM = 12.0  # 除第一个外，其余分段编号向下偏移量(mm)。
SINGLE_PRINT_SEG_NUMBER_SHIFT_LEFT_MM = 2.0   # 分段编号整体向左偏移量(mm)。
SINGLE_MARGIN_L = 2.5                       # 单拼段左边距（mm），避免刀线太靠近边缘，影响美观和实用，也避免裁剪时误伤内容。
SINGLE_MARGIN_R = 0.0                       # 单拼最右端直接贴齐 PDF 右边缘；最右图边界和刀线一起对齐到纸边。
SINGLE_SPLIT_X = 320.0                      # 单拼左右分栏的第一条分割参考线 X 坐标(mm)。
SINGLE_SPLIT_GAP_W = 10.0                   # 单拼段间隔（mm），避免刀线太靠近内容，影响美观和实用，也避免裁剪时误伤内容，同时也能让用户更清晰地看到分段。
QR_BAND = 10.0                              # QR 区高度（mm），从单拼段顶部开始算起, 这个区域内不画刀线刻度，且优先放 QR 码（如果有的话），避免和内容重叠。
QR_W = 10.0                                 # QR 码宽度（mm），放在单拼段右上角，和右边距一起占满右侧区域，避免和内容重叠。
QR_H = 10.0                                 # QR 码高度（mm），放在单拼段右上角，和 QR_BAND 一起占满顶部区域，避免和内容重叠。
SINGLE_QR_IMAGE_GAP_MM = 3.0                # 单拼中二维码下方给图片额外留白，避免首排内容贴着二维码。
SINGLE_QR_IMAGE_GAP_X_MM = 10.0             # 单拼中二维码与相邻图形之间的左右距离。
MIX_LABEL_TOP_GAP_MM = 5.0                  # 混拼中标签上方额外预留 5mm，方便裁剪。
MIX_LABEL_IMAGE_GAP_MM = 5.0                # 混拼中标签下方到图片之间额外预留 5mm，方便裁剪。
MIX_QR_TOP_GAP_MM = 5.0                     # 混拼中二维码上方额外预留 5mm，方便裁剪。
MIX_QR_IMAGE_GAP_MM = 10.0                  # 混拼中二维码/标签带与图片之间额外预留 5mm，方便裁剪。
QR_LABEL_GAP_MM = 1.5                       # 标签文本和右侧二维码之间预留空隙，避免文字挤进二维码区域。
LABEL_FONT_SIZE = 10                        # 标签字体大小（pt），放在单拼段右上角，和 QR 码一起占满右侧区域，避免和内容重叠。
RENDER_DPI = 600                            # PDF渲染分辨率（DPI），600 DPI 可以让毫米级的细节在图像上有足够的像素表现，便于后续的分析和裁剪。
CROP_PAD_MM = 1.2                           # 裁剪边界额外扩展（mm），避免裁剪时漏掉边缘内容，同时也能让成品看起来更自然一些。
TEXT_MASK_PAD_MM = 1.2                      # 自动识别外轮廓前，给文字遮罩额外扩展的宽度(mm)。
DRAW_PART_OUTER_BOX = False                 # 调试开关；True 时会额外画出每个单元的外框。
PDF_CMYK_BLACK = (0, 0, 0, 1)               # PDF 中使用的纯黑 CMYK 颜色。
PDF_CMYK_WHITE = (0, 0, 0, 0)               # PDF 中使用的白色 CMYK 颜色。
CELL_GAP_MM = 10.0                          # 混拼单元上下堆叠时的默认间距(mm)。
MIX_GAP_X_MM = 6.0                          # 混拼单元左右间距(mm)。
MARGIN_TOP_MM = 10.0                        # 混拼页面顶部内容起始留白(mm)。
MARGIN_BOTTOM_MM = 5.0                      # 混拼页面底部保底留白(mm)。
SAFE_PAD_MM = 3.0                           # 混拼页面最终高度再额外加的安全量(mm)。
PAGE_W = 600.0                              # 混拼页宽(mm)。
PAGE_H_MAX = 1500.0                         # 混拼页高上限(mm)。
PAGE_MARGIN_L = 5.0                         # 混拼左边距(mm)。
PAGE_MARGIN_R = 5.0                         # 混拼右边距(mm)。
SPLIT_X = 320.0                             # 混拼左右两栏分割线的 X 坐标(mm)。
SPLIT_GAP_W = 10.0                          # 混拼中间分栏缝宽(mm)。
MIX_THRESHOLD = 10                          # 数量小于该值时进入混拼池，数量大于等于该值时优先走单拼。
LEFT_BLOCK_W = float(SPLIT_X - PAGE_MARGIN_L)
RIGHT_BLOCK_W = float((PAGE_W - PAGE_MARGIN_R) - (SPLIT_X + SPLIT_GAP_W))
def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)
def ensure_dir(p):
    if p:
        os.makedirs(p, exist_ok=True)
def mm_to_pt(mm):
    return mm * 72.0 / 25.4
def mm_to_px(mm, dpi):
    return int(round(mm * dpi / 25.4))

def _single_print_content_w_mm():
    return max(1.0, float(SINGLE_PRINT_PAGE_W_MM) - 2.0 * float(SINGLE_PRINT_SIDE_MARGIN_MM))

def _single_print_content_h_max_mm():
    return max(1.0, float(SINGLE_H_MAX) - float(SINGLE_PRINT_TOP_MARGIN_MM) - float(SINGLE_PRINT_BOTTOM_MARGIN_MM))
def _log(log_cb, s):
    (log_cb or print)(s)
def clamp_bbox(x0, y0, x1, y1, img_w, img_h):
    x0 = max(0, min(img_w - 1, int(round(x0))))
    y0 = max(0, min(img_h - 1, int(round(y0))))
    x1 = max(x0 + 1, min(img_w, int(round(x1))))
    y1 = max(y0 + 1, min(img_h, int(round(y1))))
    return x0, y0, x1, y1
def union_bbox(b1, b2):
    if b1 is None or b2 is None:
        return b2 if b1 is None else b1
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
    return fitz.Rect(
        min(r.x0 for r in rects), min(r.y0 for r in rects),
        max(r.x1 for r in rects), max(r.y1 for r in rects)
    )
def _rect_gap_1d(a0, a1, b0, b1):
    return max(0.0, max(float(b0) - float(a1), float(a0) - float(b1)))
def _rect_close_or_overlap(a, b, gap_pt):
    if a is None or b is None:
        return False
    gx = _rect_gap_1d(a.x0, a.x1, b.x0, b.x1)
    gy = _rect_gap_1d(a.y0, a.y1, b.y0, b.y1)
    return gx <= float(gap_pt) and gy <= float(gap_pt)
def _rect_cluster_union(rects, join_gap_pt=10.0):
    rects = [fitz.Rect(r) for r in rects if r is not None and r.width > 0 and r.height > 0]
    if not rects:
        return None
    rects.sort(key=lambda r: r.width * r.height, reverse=True)
    seed = fitz.Rect(rects[0])
    seed_area = max(1.0, seed.width * seed.height)
    merged = fitz.Rect(seed)
    cluster = [seed]
    changed = True
    while changed:
        changed = False
        for r in rects[1:]:
            if any(abs(r.x0 - c.x0) <= 1e-6 and abs(r.y0 - c.y0) <= 1e-6 and abs(r.x1 - c.x1) <= 1e-6 and abs(r.y1 - c.y1) <= 1e-6 for c in cluster):
                continue
            area = max(1.0, r.width * r.height)
            if _rect_close_or_overlap(merged, r, join_gap_pt) or min(area, seed_area) >= 0.45 * max(area, seed_area):
                cluster.append(r)
                merged = rect_union([merged, r])
                changed = True
    return merged
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
def _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0, x0_mm=0.0):
    xpt = mm_to_pt(float(x0_mm))
    ypt = mm_to_pt(float(y_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, ypt), fitz.Point(xpt + tick, ypt), color=PDF_CMYK_BLACK, width=lw)
def _draw_top_edge_tick(page, x_mm, tick_len_mm=MARK_LEN, lw=1.0):
    xpt = mm_to_pt(float(x_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, tick), color=PDF_CMYK_BLACK, width=lw)
def _draw_top_tick_at_y(page, x_mm, y_mm, tick_len_mm=MARK_LEN, lw=1.0):
    xpt = mm_to_pt(float(x_mm))
    ypt = mm_to_pt(float(y_mm))
    tick = mm_to_pt(float(tick_len_mm))
    page.draw_line(fitz.Point(xpt, ypt), fitz.Point(xpt, ypt + tick), color=PDF_CMYK_BLACK, width=lw)
EDGE_SNAP_MM = 10.0   # <=6mm 就吸到纸边（你现在通常是 5mm）
def _snap_edge_mm(v, lo, hi, eps):
    if abs(v - lo) <= eps:
        return lo
    if abs(v - hi) <= eps:
        return hi
    return v
def _draw_segment_cut_ticks(page, seg, yoff_mm, page_h_mm=SINGLE_H_MAX, draw_left_ticks=True):
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
    W = _single_print_content_w_mm()
    x_min = float(min(x_lines))
    x_max = float(max(x_lines))
    draw_xs = set()
    for x in x_lines:
        x_raw = float(x)
        x_draw = x_raw
        if abs(x - x_min) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        elif abs(x - x_max) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        if 0.0 - 1e-6 <= x_raw <= W + 1e-6:
            draw_xs.add(round(x_raw, 3))
        if 0.0 - 1e-6 <= x_draw <= W + 1e-6:
            draw_xs.add(round(x_draw, 3))
    for x_draw in sorted(draw_xs):
        _draw_top_tick_at_y(page, x_draw, yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)
    y_lines = set(ys)
    for y in ys:
        y_lines.add(round(y + h, 3))
    y_lines = sorted(y_lines)
    if not y_lines:
        return
    y_min = float(min(y_lines))
    y_max = float(max(y_lines))
    if not draw_left_ticks:
        return
    for y in y_lines:
        Y = float(yoff_mm) + float(y)
        if abs(y - y_max) <= 1e-6:
            Y = _snap_edge_mm(Y, 0.0, float(page_h_mm), EDGE_SNAP_MM)
        if 0.0 - 1e-6 <= Y <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, Y, tick_len_mm=MARK_LEN, lw=1.0)

def _redraw_single_top_ticks(page, seg, yoff_mm, exclude_x_range=None):
    pls = seg.get("placements") or []
    if not pls:
        return
    w = float(pls[0].get("w", 0.0))
    if w <= 0:
        return
    xs = sorted({round(float(p["x"]), 3) for p in pls})
    if not xs:
        return
    x_lines = set(xs)
    for x in xs:
        x_lines.add(round(x + w, 3))
    W = _single_print_content_w_mm()
    x_min = float(min(x_lines))
    x_max = float(max(x_lines))
    lo = hi = None
    if exclude_x_range is not None:
        lo, hi = exclude_x_range
        lo = float(lo)
        hi = float(hi)
    draw_xs = set()
    for x in x_lines:
        x_raw = float(x)
        x_draw = x_raw
        if abs(x - x_min) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        elif abs(x - x_max) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        for xv in (x_raw, x_draw):
            if not (0.0 - 1e-6 <= xv <= W + 1e-6):
                continue
            if lo is not None and hi is not None and (lo + 1e-6) < xv < (hi - 1e-6):
                continue
            draw_xs.add(round(xv, 3))
    for x_draw in sorted(draw_xs):
        _draw_top_tick_at_y(page, x_draw, yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)
def _is_single_special_mode(seg=None):
    if not bool(SINGLE_SINGLE_LINE_SPECIAL_MODE):
        return False
    if seg is None:
        return False
    return not bool((seg or {}).get("single_is_double_cut", True))
def _single_special_tick_positions(seg, yoff_mm):
    pls = list((seg or {}).get("placements") or [])
    if not pls:
        return [], []
    x_intervals = [
        (float(p["x"]), float(p["x"]) + float(p["w"]))
        for p in pls
    ]
    y_intervals = [
        (float(yoff_mm) + float(p["y"]), float(yoff_mm) + float(p["y"]) + float(p["h"]))
        for p in pls
    ]
    x_blocks = _merge_intervals(x_intervals, tol=1e-3)
    y_blocks = _merge_intervals(y_intervals, tol=1e-3)
    top_xs = set()
    left_ys = set()
    if x_blocks:
        top_xs.add(float(x_blocks[0][0]))
        top_xs.add(float(x_blocks[-1][1]))
        for (_ax0, ax1), (bx0, _bx1) in zip(x_blocks, x_blocks[1:]):
            top_xs.add((float(ax1) + float(bx0)) / 2.0)
    if y_blocks:
        left_ys.add(float(y_blocks[0][0]))
        left_ys.add(float(y_blocks[-1][1]))
        for (_ay0, ay1), (by0, _by1) in zip(y_blocks, y_blocks[1:]):
            left_ys.add((float(ay1) + float(by0)) / 2.0)
    return sorted(top_xs), sorted(left_ys)
def _single_left_tick_ys(seg, yoff_mm):
    if _is_single_special_mode(seg):
        _top_xs, left_ys = _single_special_tick_positions(seg, yoff_mm)
        return [float(y) for y in left_ys]
    pls = list((seg or {}).get("placements") or [])
    if not pls:
        return []
    h = float(pls[0].get("h", 0.0))
    if h <= 0:
        return []
    ys = sorted({round(float(p["y"]), 3) for p in pls})
    y_lines = set(ys)
    for y in ys:
        y_lines.add(round(float(y) + float(h), 3))
    return [float(yoff_mm) + float(y) for y in sorted(y_lines)]
def _draw_single_special_edge_ticks(page, seg, yoff_mm, page_h_mm, exclude_x_range=None, top_only=False):
    top_xs, left_ys = _single_special_tick_positions(seg, yoff_mm)
    lo = hi = None
    if exclude_x_range is not None:
        lo, hi = exclude_x_range
        lo = float(lo)
        hi = float(hi)
    for x in top_xs:
        if lo is not None and hi is not None and (lo + 1e-6) < float(x) < (hi - 1e-6):
            continue
        if 0.0 - 1e-6 <= float(x) <= _single_print_content_w_mm() + 1e-6:
            _draw_top_tick_at_y(page, float(x), yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)
    if top_only:
        return
    for y in left_ys:
        if 0.0 - 1e-6 <= float(y) <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, float(y), tick_len_mm=MARK_LEN, lw=1.0)
def _draw_single_outer_left_cut_ticks(page, sheet, outer_h_mm):
    x0_mm = max(0.0, float(SINGLE_PRINT_SIDE_MARGIN_MM) - float(SINGLE_LEFT_CUT_TICK_SHIFT_LEFT_MM))
    yoff = 0.0
    for idx, seg in enumerate(sheet.get("segments") or []):
        if idx > 0:
            y_mm = float(SINGLE_PRINT_TOP_MARGIN_MM) + float(yoff)
            if 0.0 - 1e-6 <= y_mm <= float(outer_h_mm) + 1e-6:
                _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0, x0_mm=x0_mm)
        for y in _single_left_tick_ys(seg, yoff):
            y_mm = float(SINGLE_PRINT_TOP_MARGIN_MM) + float(y)
            if 0.0 - 1e-6 <= y_mm <= float(outer_h_mm) + 1e-6:
                _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0, x0_mm=x0_mm)
        yoff += float((seg.get("best") or {}).get("H", 0.0))

def _single_top_tick_xs(seg, yoff_mm=0.0):
    if _is_single_special_mode(seg):
        top_xs, _left_ys = _single_special_tick_positions(seg, yoff_mm)
        return [float(x) for x in top_xs]
    pls = list((seg or {}).get("placements") or [])
    if not pls:
        return []
    w = float(pls[0].get("w", 0.0))
    if w <= 0:
        return []
    xs = sorted({round(float(p["x"]), 3) for p in pls})
    x_lines = set(xs)
    for x in xs:
        x_lines.add(round(float(x) + float(w), 3))
    if not x_lines:
        return []
    W = _single_print_content_w_mm()
    x_min = float(min(x_lines))
    x_max = float(max(x_lines))
    draw_xs = set()
    for x in x_lines:
        x_raw = float(x)
        x_draw = x_raw
        if abs(float(x) - x_min) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        elif abs(float(x) - x_max) <= 1e-6:
            x_draw = _snap_edge_mm(x_draw, 0.0, W, EDGE_SNAP_MM)
        if 0.0 - 1e-6 <= x_raw <= W + 1e-6:
            draw_xs.add(float(x_raw))
        if 0.0 - 1e-6 <= x_draw <= W + 1e-6:
            draw_xs.add(float(x_draw))
    return sorted(draw_xs)

def _single_qr_rect_mm(seg, yoff_mm):
    content_w = _single_print_content_w_mm()
    qr_x0 = float(content_w - QR_W)
    placements = list(seg.get("placements") or [])
    if placements:
        img_right = max(float(p.get("x", 0.0)) + float(p.get("w", 0.0)) for p in placements)
        qr_x0 = min(float(content_w - QR_W), float(img_right) + float(SINGLE_QR_IMAGE_GAP_X_MM))
    if _is_single_special_mode(seg):
        qr_x0 -= float(SINGLE_SPECIAL_QR_SHIFT_LEFT_MM)
    top_tick_xs = _single_top_tick_xs(seg, yoff_mm=yoff_mm)
    if top_tick_xs:
        right_cut_x = float(max(top_tick_xs))
        if float(qr_x0) <= right_cut_x <= float(qr_x0 + QR_W):
            qr_x0 = min(float(qr_x0), float(right_cut_x) - float(SINGLE_QR_RIGHT_CUTLINE_GAP_MM) - float(QR_W))
    qr_x0 = max(0.0, min(float(content_w - QR_W), float(qr_x0)))
    return (
        float(qr_x0),
        float(yoff_mm),
        float(qr_x0 + QR_W),
        float(yoff_mm + QR_H),
    )
def _draw_horizontal_cutline(page, y_mm, lw=1.0):
    ypt = mm_to_pt(float(y_mm))
    Wpt = mm_to_pt(_single_print_content_w_mm())
    page.draw_line(fitz.Point(0, ypt), fitz.Point(Wpt, ypt), color=PDF_CMYK_BLACK, width=lw)
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

def _format_number_token(v):
    try:
        fv = float(v)
    except Exception:
        s = str(v).strip()
        return s
    if abs(fv - round(fv)) <= 1e-9:
        return str(int(round(fv)))
    s = ("%.6f" % fv).rstrip("0").rstrip(".")
    return s

def _normalize_size_label(size_text):
    s = str(size_text or "").strip()
    m = re.search(r"(\d+(\.\d+)?)\s*[\*xX×✕]\s*(\d+(\.\d+)?)", s)
    if m:
        return "%sX%s" % (_format_number_token(m.group(1)), _format_number_token(m.group(3)))
    s = re.sub(r"\s*[\*xX×✕]\s*", "X", s)
    return s.replace(" ", "").upper()
def _extract_order_meta(pdf_path):
    base = os.path.splitext(os.path.basename(str(pdf_path or "")))[0]
    ps = base.split("^")
    if len(ps) < 11:
        return None
    qty_match = re.search(r"(\d+)", str(ps[7]))
    return {
        "base": base,
        "parts": ps,
        "customer_id": str(ps[1]).strip(),
        "size_text": _normalize_size_label(ps[6]),
        "qty_text": qty_match.group(1) if qty_match else str(ps[7]).strip(),
        "order_id": str(ps[9]).strip(),
        "serial_id": str(ps[10]).strip(),
    }
def extract_label_text(pdf_path):
    meta = _extract_order_meta(pdf_path)
    if meta:
        customer_id = meta["customer_id"]
        size_text = meta["size_text"]
        qty_text = meta["qty_text"]
        order_id = meta["order_id"]
        serial_id = meta["serial_id"]
        head = ""
        if order_id:
            head = "订单:" + order_id
        if serial_id:
            head = (head + "^" + serial_id) if head else serial_id

        details = []
        if customer_id:
            details.append("ID " + customer_id)
        if size_text:
            details.append("单个规格" + size_text)
        if qty_text:
            details.append("数量" + qty_text + "枚")

        if head and details:
            return head + "(" + "  ".join(details) + ")"
        if head:
            return head
        base = meta["base"]
        ps = meta["parts"]
    else:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        ps = base.split("^")
    if len(ps) >= 2:
        return "^".join(ps[:2])
    return base[:40]
def build_single_output_name(pdf_path):
    meta = _extract_order_meta(pdf_path)
    if not meta:
        base = os.path.splitext(os.path.basename(str(pdf_path or "")))[0]
        return base[:80] or "single"
    head_parts = [meta["order_id"], meta["serial_id"]]
    head = "^".join([part for part in head_parts if part])
    details = []
    if meta["customer_id"]:
        details.append("ID" + meta["customer_id"])
    if meta["size_text"]:
        details.append("单个规格" + meta["size_text"])
    if meta["qty_text"]:
        details.append("数量" + meta["qty_text"] + "枚")
    if head and details:
        return head + "(" + "-".join(details) + ")"
    return head or (meta["base"][:80] or "single")
def _sanitize_output_stem(name, fallback="single"):
    stem = re.sub(r'[<>:"/\\|?*]+', "_", str(name or "").strip())
    stem = stem.rstrip(" .")
    return stem or str(fallback)
def _build_single_sheet_pdf_name(sheet, sheet_index, name_counts):
    unique_names = []
    for seg in (sheet.get("segments") or []):
        name = str(seg.get("single_output_name") or "").strip()
        if name and name not in unique_names:
            unique_names.append(name)
    if not unique_names:
        base = "single_%d" % int(sheet_index)
    elif len(unique_names) == 1:
        base = unique_names[0]
    else:
        base = "%s_等%d款" % (unique_names[0], len(unique_names))
    base = _sanitize_output_stem(base, fallback="single_%d" % int(sheet_index))
    seq = int(name_counts.get(base, 0)) + 1
    name_counts[base] = seq
    if seq > 1:
        base = "%s_%d" % (base, seq)
    return base + ".pdf"
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
def get_single_layout_rule_from_filename(pdf_path):
    base = os.path.splitext(os.path.basename(str(pdf_path or "")))[0]
    is_double_cut = ("\u53cc\u679a" in base) or ("\u53cc\u5207" in base)
    outer_ext_mm = float(OUTER_EXT if is_double_cut else SINGLE_CUT_OUTER_EXT_MM)
    return {
        "is_double_cut": bool(is_double_cut),
        "outer_ext_mm": float(outer_ext_mm),
        "inner_margin_mm": float(INNER_MARGIN_MM if is_double_cut else outer_ext_mm),
    }
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
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsunb.ttf",
        r"C:\Windows\Fonts\simfang.ttf",
        r"C:\Windows\Fonts\simkai.ttf",
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simsun.ttc",
    ]
    return next((p for p in candidates if os.path.exists(p)), None)
def pick_arial_bold_italic_fontfile():
    candidates = [
        r"C:\Windows\Fonts\arialbi.ttf",
        r"C:\Windows\Fonts\Arialbi.ttf",
        r"C:\Windows\Fonts\ariali.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    return next((p for p in candidates if os.path.exists(p)), None)
def _single_print_seg_text_kwargs(page):
    fontfile = pick_arial_bold_italic_fontfile()
    if fontfile:
        try:
            page.insert_font(fontname="F_SEG_ARIAL_BI", fontfile=fontfile)
            return {"fontname": "F_SEG_ARIAL_BI", "color": PDF_CMYK_BLACK}
        except Exception:
            return {"fontfile": fontfile, "color": PDF_CMYK_BLACK}
    return {"fontname": "helv", "color": PDF_CMYK_BLACK}
def _char_width_factor(ch):
    o = ord(ch)
    if ch in " \t":
        return 0.33
    if ch in ",.;:!|/\\'\"`":
        return 0.32
    if ch in "()[]{}<>":
        return 0.38
    if ch in "-_^":
        return 0.42
    if ch.isdigit():
        return 0.58
    if ("A" <= ch <= "Z") or ("a" <= ch <= "z"):
        return 0.62
    if 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF:
        return 1.00
    return 0.70

def _estimate_text_width_pt(text, fontsize):
    total = 0.0
    for ch in str(text or ""):
        total += _char_width_factor(ch) * float(fontsize)
    return total

def _split_chunk_to_width(chunk, fontsize, max_width_pt):
    chunk = str(chunk or "")
    if not chunk:
        return [""]
    lines = []
    cur = []
    cur_w = 0.0
    last_break = -1
    break_chars = set([" ", "(", "（", ")", "）", "^", "_", "-", "单", "数"])
    for ch in chunk:
        ch_w = _estimate_text_width_pt(ch, fontsize)
        if cur and (cur_w + ch_w) > float(max_width_pt):
            cut = len(cur)
            if last_break >= 0:
                cut = last_break + 1
            piece = "".join(cur[:cut]).rstrip()
            if piece:
                lines.append(piece)
            rest = "".join(cur[cut:]).lstrip()
            cur = list(rest) if rest else []
            cur_w = _estimate_text_width_pt(rest, fontsize)
            last_break = -1
            for idx, old_ch in enumerate(cur):
                if old_ch in break_chars:
                    last_break = idx
        cur.append(ch)
        cur_w += ch_w
        if ch in break_chars:
            last_break = len(cur) - 1
    if cur:
        lines.append("".join(cur).rstrip())
    return [line for line in lines if line]

def _wrap_text_to_width(text, fontsize, max_width_pt):
    text = str(text or "")
    if not text:
        return []
    lines = []
    for chunk in text.splitlines() or [text]:
        if not chunk:
            lines.append("")
            continue
        lines.extend(_split_chunk_to_width(chunk, fontsize, max_width_pt))
    return lines

def _fit_text_lines(text, max_width_pt, max_height_pt, base_fontsize=LABEL_FONT_SIZE, min_fontsize=4.0):
    text = str(text or "").strip()
    if not text:
        return 0.0, []
    fs = float(base_fontsize)
    while fs >= float(min_fontsize) - 1e-9:
        line_h = fs * 1.22
        max_lines = max(1, int(float(max_height_pt) / line_h))
        lines = _wrap_text_to_width(text, fs, max_width_pt)
        if len(lines) <= max_lines:
            return fs, lines
        fs -= 0.5
    fs = float(min_fontsize)
    line_h = fs * 1.22
    max_lines = max(1, int(float(max_height_pt) / line_h))
    lines = _wrap_text_to_width(text, fs, max_width_pt)
    return fs, lines[:max_lines]

def _ensure_label_font(page, fontfile):
    if not fontfile:
        return None
    fontname = "F_LABEL_CJK"
    try:
        page.insert_font(fontname=fontname, fontfile=fontfile)
        return fontname
    except Exception:
        return None
def _label_text_kwargs(page, fontfile):
    fontname = _ensure_label_font(page, fontfile)
    if fontname:
        return {"fontname": fontname, "color": PDF_CMYK_BLACK}
    if fontfile:
        return {"fontfile": fontfile, "color": PDF_CMYK_BLACK}
    return {"fontname": "helv", "color": PDF_CMYK_BLACK}

def _draw_label_lines(page, lines, xpt, ypt, fontsize, fontfile):
    if not lines:
        return
    text_kwargs = _label_text_kwargs(page, fontfile)
    line_gap = float(fontsize) * 1.22
    for idx, line in enumerate(lines):
        y_line = ypt + idx * line_gap
        page.insert_text(fitz.Point(xpt, y_line), line, fontsize=fontsize, **text_kwargs)
def _draw_qr_on_white(page, qr_rect, qr_bytes, pad_left_mm=0.0, pad_top_mm=0.0, pad_right_mm=0.0, pad_bottom_mm=0.0):
    if not qr_bytes:
        return
    pad_left_pt = mm_to_pt(float(pad_left_mm))
    pad_top_pt = mm_to_pt(float(pad_top_mm))
    pad_right_pt = mm_to_pt(float(pad_right_mm))
    pad_bottom_pt = mm_to_pt(float(pad_bottom_mm))
    bg_rect = fitz.Rect(
        qr_rect.x0 - pad_left_pt,
        qr_rect.y0 - pad_top_pt,
        qr_rect.x1 + pad_right_pt,
        qr_rect.y1 + pad_bottom_pt,
    )
    page.draw_rect(bg_rect, color=PDF_CMYK_WHITE, fill=PDF_CMYK_WHITE, width=0)
    page.insert_image(qr_rect, stream=qr_bytes, keep_proportion=True)
def _prepare_mix_label_text(text):
    s = str(text or "").replace("\r", "\n").strip()
    if not s:
        return ""
    return s
def _mix_label_available_w_mm(block_w_mm):
    return max(1.0, float(block_w_mm) - float(QR_W) - float(QR_LABEL_GAP_MM) - 1.0)
def _mix_min_marker_band_mm():
    return float(max(
        float(QR_BAND),
        float(MIX_QR_TOP_GAP_MM) + float(QR_H) + float(MIX_QR_IMAGE_GAP_MM),
    ))
def _estimate_mix_marker_band_mm(text, block_w_mm, min_band_mm=QR_BAND):
    prepared = _prepare_mix_label_text(text)
    if not prepared:
        return float(max(0.0, float(min_band_mm)))
    max_w_pt = mm_to_pt(max(1.0, _mix_label_available_w_mm(block_w_mm) - 1.2))
    fs = 4.0
    lines = _wrap_text_to_width(prepared, fs, max_w_pt)
    text_h_pt = max(1.0, float(len(lines)) * fs * 1.22)
    text_h_mm = text_h_pt * 25.4 / 72.0
    label_need_mm = float(MIX_LABEL_TOP_GAP_MM) + text_h_mm + float(MIX_LABEL_IMAGE_GAP_MM)
    return float(max(float(min_band_mm), label_need_mm))
def _mix_marker_band_for_width(t, block_w_mm, marker_keys=None):
    if _mix_marker_key(t) in set(marker_keys or []):
        return 0.0
    need_qr = bool((t or {}).get("qr_bytes"))
    need_label = bool(str((t or {}).get("label_text") or "").strip())
    if not (need_qr or need_label):
        return 0.0
    min_band_mm = _mix_min_marker_band_mm() if need_qr else 0.0
    return _estimate_mix_marker_band_mm((t or {}).get("label_text", ""), block_w_mm, min_band_mm=min_band_mm)
def _mix_make_block(t, x0, x1, y0, y1, marker_band, image_band):
    marker_band = float(marker_band)
    return {
        "type": t, "x0": float(x0), "x1": float(x1), "y0": float(y0), "y1": float(y1),
        "marker_band": marker_band, "image_band": float(image_band), "needs_marker": bool(marker_band > 1e-9),
    }
def _mix_make_placement(t, x, y, w, h, rot90, marker_band, image_band):
    return {
        "type": t, "x": float(x), "y": float(y), "w": float(w), "h": float(h),
        "rot90": bool(rot90), "marker_band": float(marker_band), "image_band": float(image_band),
    }

def _draw_label_top_left(page, text, x_left_mm, y_top_mm, max_w_mm, band_h_mm):
    fontfile = pick_cjk_fontfile()
    prepared = _prepare_mix_label_text(text)
    max_w_pt = mm_to_pt(max(1.0, float(max_w_mm) - 1.2))
    max_h_pt = mm_to_pt(max(3.0, float(band_h_mm) - float(MIX_LABEL_TOP_GAP_MM) - float(MIX_LABEL_IMAGE_GAP_MM) - 0.6))
    fontsize, lines = _fit_text_lines(prepared, max_w_pt, max_h_pt)
    if not lines:
        return
    text_kwargs = _label_text_kwargs(page, fontfile)
    x_left_pt = mm_to_pt(float(x_left_mm) + 0.5)
    ypt = mm_to_pt(float(y_top_mm) + float(MIX_LABEL_TOP_GAP_MM) + 0.3) + float(fontsize)
    line_gap = float(fontsize) * 1.22
    for idx, line in enumerate(lines):
        y_line = ypt + idx * line_gap
        line_w = _estimate_text_width_pt(line, fontsize)
        xpt = x_left_pt + max(0.0, (float(max_w_pt) - float(line_w)) / 2.0)
        page.insert_text(fitz.Point(xpt, y_line), line, fontsize=fontsize, **text_kwargs)
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
        "image_bbox": _rect_cluster_union(image_boxes),
        "draw_bbox": _rect_cluster_union(draw_rects),
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
def _bbox_area(bbox):
    if bbox is None:
        return -1.0
    return max(0.0, float(bbox[2] - bbox[0])) * max(0.0, float(bbox[3] - bbox[1]))
def _bbox_gap_1d(a0, a1, b0, b1):
    return max(0.0, max(float(b0) - float(a1), float(a0) - float(b1)))
def _bbox_close_or_overlap(a, b, gap_px):
    if a is None or b is None:
        return False
    gx = _bbox_gap_1d(a[0], a[2], b[0], b[2])
    gy = _bbox_gap_1d(a[1], a[3], b[1], b[3])
    return gx <= float(gap_px) and gy <= float(gap_px)
def _bbox_contains(a, b, eps=1.0):
    if a is None or b is None:
        return False
    return (
        float(a[0]) <= float(b[0]) + float(eps) and
        float(a[1]) <= float(b[1]) + float(eps) and
        float(a[2]) >= float(b[2]) - float(eps) and
        float(a[3]) >= float(b[3]) - float(eps)
    )
def _merge_bbox_candidates(img_w, img_h, *bboxes):
    valid = []
    fallback = None
    for bbox in bboxes:
        if bbox is None:
            continue
        if fallback is None:
            fallback = bbox
        good = _meaningful_bbox_or_none(bbox, img_w, img_h)
        if good is not None:
            valid.append(good)
    if not valid:
        return fallback
    valid.sort(key=_bbox_area)
    best = valid[0]
    best_area = max(1.0, _bbox_area(best))
    join_gap_px = max(18.0, 0.01 * float(min(img_w, img_h)))
    for bbox in valid[1:]:
        area = max(1.0, _bbox_area(bbox))
        if _bbox_contains(best, bbox):
            continue
        if _bbox_contains(bbox, best):
            if area <= best_area * 1.25:
                best = bbox
                best_area = area
            continue
        if _bbox_close_or_overlap(best, bbox, join_gap_px) or min(area, best_area) >= 0.45 * max(area, best_area):
            best = union_bbox(best, bbox)
            best_area = max(1.0, _bbox_area(best))
    return best
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
            keep.append((float(area), (x, y, x + w, y + h)))
    if not keep:
        _, x, y, w, h = bboxes[0]
        return (x, y, x + w, y + h)
    seed_area, seed_bbox = keep[0]
    cluster = [seed_bbox]
    merged = seed_bbox
    join_gap_px = max(18.0, 0.012 * float(min(W, H)))
    changed = True
    while changed:
        changed = False
        for area, bbox in keep[1:]:
            if bbox in cluster:
                continue
            if float(area) < 0.18 * float(seed_area) and not _bbox_close_or_overlap(merged, bbox, join_gap_px):
                continue
            if _bbox_close_or_overlap(merged, bbox, join_gap_px):
                cluster.append(bbox)
                merged = union_bbox(merged, bbox)
                changed = True
    return tuple(int(round(v)) for v in merged)
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
    merged = _merge_bbox_candidates(W, H, *candidates)
    return merged if merged is not None else candidates[0]
def get_page_bbox_candidates_px(doc, page_index, pil_img):
    info = get_page_struct_info(doc, page_index)
    page_rect = info["page_rect"]
    text_boxes = info["text_boxes"]
    img_w, img_h = pil_img.size
    draw_bbox_px = pdf_rect_to_px_bbox(info["draw_bbox"], page_rect, img_w, img_h)
    image_bbox_px = pdf_rect_to_px_bbox(info["image_bbox"], page_rect, img_w, img_h)
    masked = mask_text_regions_on_pil(pil_img, text_boxes, page_rect, pad_mm=TEXT_MASK_PAD_MM)
    cv_bbox_masked = find_outer_bbox(masked)
    cv_bbox_raw = find_outer_bbox(pil_img)
    cv_bbox_px = _merge_bbox_candidates(img_w, img_h, cv_bbox_masked, cv_bbox_raw)
    return draw_bbox_px, image_bbox_px, cv_bbox_px
def _meaningful_bbox_or_none(bbox, img_w, img_h):
    if bbox is None:
        return None
    x0, y0, x1, y1 = bbox
    bw = int(x1) - int(x0)
    bh = int(y1) - int(y0)
    if bw < 8 or bh < 8:
        return None
    if bw >= 0.995 * float(img_w) and bh >= 0.995 * float(img_h):
        return None
    return clamp_bbox(x0, y0, x1, y1, img_w, img_h)
def _scale_bbox_between_sizes(bbox, src_size, dst_size):
    if bbox is None:
        return None
    sw, sh = src_size
    dw, dh = dst_size
    if sw <= 0 or sh <= 0 or dw <= 0 or dh <= 0:
        return None
    x0, y0, x1, y1 = bbox
    sx = float(dw) / float(sw)
    sy = float(dh) / float(sh)
    return (
        int(round(x0 * sx)),
        int(round(y0 * sy)),
        int(round(x1 * sx)),
        int(round(y1 * sy)),
    )
def _bbox_with_pad_or_full(bbox, img_size, pad_px=0):
    img_w, img_h = img_size
    if bbox is None:
        return (0, 0, int(img_w), int(img_h))
    x0, y0, x1, y1 = bbox
    return clamp_bbox(x0 - pad_px, y0 - pad_px, x1 + pad_px, y1 + pad_px, img_w, img_h)
def _px_bbox_to_pdf_rect(bbox, page_rect_pt, img_size):
    if bbox is None or page_rect_pt is None:
        return None
    img_w, img_h = img_size
    if img_w <= 0 or img_h <= 0:
        return None
    x0, y0, x1, y1 = bbox
    sx = float(page_rect_pt.width) / float(img_w)
    sy = float(page_rect_pt.height) / float(img_h)
    return fitz.Rect(
        float(page_rect_pt.x0) + float(x0) * sx,
        float(page_rect_pt.y0) + float(y0) * sy,
        float(page_rect_pt.x0) + float(x1) * sx,
        float(page_rect_pt.y0) + float(y1) * sy,
    )
def _build_src_spec_from_px_bbox(pdf_path, page_index, bbox_px, page_rect_pt, img_size):
    clip_rect = _px_bbox_to_pdf_rect(bbox_px, page_rect_pt, img_size)
    if clip_rect is None or clip_rect.width <= 0 or clip_rect.height <= 0:
        return None
    return {
        "pdf_path": str(pdf_path),
        "page_index": int(page_index),
        "clip_rect": [float(clip_rect.x0), float(clip_rect.y0), float(clip_rect.x1), float(clip_rect.y1)],
    }
def _show_pdf_clip(page, target_rect, src_spec, src_doc_cache, rotate_deg=0, keep_proportion=True):
    if not src_spec:
        return False
    pdf_path = str(src_spec.get("pdf_path") or "")
    page_index = int(src_spec.get("page_index", 0))
    clip_rect_vals = src_spec.get("clip_rect") or []
    if (not pdf_path) or len(clip_rect_vals) != 4:
        return False
    try:
        src_doc = None
        cache = src_doc_cache if isinstance(src_doc_cache, dict) else {}
        current_path = str(cache.get("__current_path__") or "")
        current_doc = cache.get("__current_doc__")
        if current_path == pdf_path and current_doc is not None:
            src_doc = current_doc
        else:
            if current_doc is not None:
                try:
                    current_doc.close()
                except Exception:
                    pass
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            src_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            cache["__current_path__"] = pdf_path
            cache["__current_doc__"] = src_doc
            cache["__current_bytes__"] = pdf_bytes
        if page_index < 0 or page_index >= src_doc.page_count:
            return False
        clip_rect = fitz.Rect(*clip_rect_vals)
        page.show_pdf_page(
            target_rect,
            src_doc,
            page_index,
            clip=clip_rect,
            rotate=int(rotate_deg),
            keep_proportion=bool(keep_proportion),
        )
        return True
    except Exception:
        return False
def _close_doc_cache(doc_cache):
    seen = set()
    for doc in list((doc_cache or {}).values()):
        if doc is None:
            continue
        doc_id = id(doc)
        if doc_id in seen:
            continue
        seen.add(doc_id)
        try:
            doc.close()
        except Exception:
            pass
def _content_bbox_from_png_bytes(img_bytes, white_thr=250):
    if not img_bytes:
        return None, None
    try:
        with Image.open(BytesIO(img_bytes)) as im:
            rgb = im.convert("RGB")
            mask = rgb.convert("L").point(lambda v: 255 if v < int(white_thr) else 0)
            return mask.getbbox(), tuple(rgb.size)
    except Exception:
        return None, None
def _trim_src_spec_clip_by_png_content(src_spec, img_bytes):
    if not src_spec or not img_bytes:
        return src_spec
    clip_rect_vals = list(src_spec.get("clip_rect") or [])
    if len(clip_rect_vals) != 4:
        return src_spec
    bbox, img_size = _content_bbox_from_png_bytes(img_bytes)
    if bbox is None or not img_size:
        return src_spec
    img_w, img_h = img_size
    if img_w <= 0 or img_h <= 0:
        return src_spec
    bx0, by0, bx1, by1 = bbox
    clip_x0, clip_y0, clip_x1, clip_y1 = [float(v) for v in clip_rect_vals]
    clip_w = clip_x1 - clip_x0
    clip_h = clip_y1 - clip_y0
    if clip_w <= 0 or clip_h <= 0:
        return src_spec
    left_pt = clip_w * float(bx0) / float(img_w)
    top_pt = clip_h * float(by0) / float(img_h)
    right_pt = clip_w * float(max(0, img_w - bx1)) / float(img_w)
    bottom_pt = clip_h * float(max(0, img_h - by1)) / float(img_h)
    new_x0 = clip_x0 + left_pt
    new_y0 = clip_y0 + top_pt
    new_x1 = clip_x1 - right_pt
    new_y1 = clip_y1 - bottom_pt
    if new_x1 <= new_x0 + 1e-6 or new_y1 <= new_y0 + 1e-6:
        return src_spec
    new_spec = dict(src_spec)
    new_spec["clip_rect"] = [new_x0, new_y0, new_x1, new_y1]
    return new_spec
def _crop_png_bytes_to_content(img_bytes):
    bbox, img_size = _content_bbox_from_png_bytes(img_bytes)
    if bbox is None or not img_size:
        return img_bytes
    try:
        with Image.open(BytesIO(img_bytes)) as im:
            cropped = im.crop(bbox).convert("RGBA")
        return _pil_to_png_bytes(cropped)
    except Exception:
        return img_bytes
def _trim_src_spec_edges_mm(src_spec, trim_left_mm=0.0, trim_top_mm=0.0, trim_right_mm=0.0, trim_bottom_mm=0.0):
    if not src_spec:
        return src_spec
    clip_rect_vals = list(src_spec.get("clip_rect") or [])
    if len(clip_rect_vals) != 4:
        return src_spec
    x0, y0, x1, y1 = [float(v) for v in clip_rect_vals]
    left_pt = mm_to_pt(float(max(0.0, trim_left_mm)))
    top_pt = mm_to_pt(float(max(0.0, trim_top_mm)))
    right_pt = mm_to_pt(float(max(0.0, trim_right_mm)))
    bottom_pt = mm_to_pt(float(max(0.0, trim_bottom_mm)))
    new_x0 = x0 + left_pt
    new_y0 = y0 + top_pt
    new_x1 = x1 - right_pt
    new_y1 = y1 - bottom_pt
    if new_x1 <= new_x0 + 1e-6 or new_y1 <= new_y0 + 1e-6:
        return src_spec
    new_spec = dict(src_spec)
    new_spec["clip_rect"] = [new_x0, new_y0, new_x1, new_y1]
    return new_spec
def _crop_png_bytes_by_src_trim(img_bytes, src_spec, trim_left_mm=0.0, trim_top_mm=0.0, trim_right_mm=0.0, trim_bottom_mm=0.0):
    if (not img_bytes) or (not src_spec):
        return img_bytes
    clip_rect_vals = list(src_spec.get("clip_rect") or [])
    if len(clip_rect_vals) != 4:
        return img_bytes
    try:
        with Image.open(BytesIO(img_bytes)) as im:
            rgba = im.convert("RGBA")
            img_w, img_h = rgba.size
            clip_x0, clip_y0, clip_x1, clip_y1 = [float(v) for v in clip_rect_vals]
            clip_w_pt = max(1e-6, clip_x1 - clip_x0)
            clip_h_pt = max(1e-6, clip_y1 - clip_y0)
            left_px = int(round(mm_to_pt(float(max(0.0, trim_left_mm))) * float(img_w) / clip_w_pt))
            top_px = int(round(mm_to_pt(float(max(0.0, trim_top_mm))) * float(img_h) / clip_h_pt))
            right_px = int(round(mm_to_pt(float(max(0.0, trim_right_mm))) * float(img_w) / clip_w_pt))
            bottom_px = int(round(mm_to_pt(float(max(0.0, trim_bottom_mm))) * float(img_h) / clip_h_pt))
            x0 = max(0, min(img_w - 1, left_px))
            y0 = max(0, min(img_h - 1, top_px))
            x1 = max(x0 + 1, min(img_w, img_w - right_px))
            y1 = max(y0 + 1, min(img_h, img_h - bottom_px))
            cropped = rgba.crop((x0, y0, x1, y1))
        return _pil_to_png_bytes(cropped)
    except Exception:
        return img_bytes
def _pil_to_png_bytes(pil_img):
    bio = BytesIO()
    pil_img.save(bio, format="PNG", optimize=True)
    return bio.getvalue()
def _rotate_png_bytes_if_needed(img_bytes, rot90):
    if (not rot90) or (not img_bytes):
        return img_bytes
    try:
        with Image.open(BytesIO(img_bytes)) as im:
            rotated = im.transpose(Image.ROTATE_90).convert("RGBA")
        return _pil_to_png_bytes(rotated)
    except Exception:
        return img_bytes
def _detect_shared_pair_bbox(doc, page_body, page_cont, img_body, img_cont):
    body_draw, body_img, body_cv = get_page_bbox_candidates_px(doc, page_body, img_body)
    cont_draw, cont_img, cont_cv = get_page_bbox_candidates_px(doc, page_cont, img_cont)
    body_bbox = _merge_bbox_candidates(img_body.size[0], img_body.size[1], body_cv, body_draw, body_img)
    cont_bbox = _merge_bbox_candidates(img_cont.size[0], img_cont.size[1], cont_cv, cont_draw, cont_img)
    body_bbox = _meaningful_bbox_or_none(body_bbox, *img_body.size)
    cont_bbox = _meaningful_bbox_or_none(cont_bbox, *img_cont.size)
    if body_bbox is None and cont_bbox is None:
        return None
    if img_body.size == img_cont.size:
        return _merge_bbox_candidates(img_body.size[0], img_body.size[1], body_bbox, cont_bbox)
    scaled_cont = _scale_bbox_between_sizes(cont_bbox, img_cont.size, img_body.size)
    scaled_cont = _meaningful_bbox_or_none(scaled_cont, *img_body.size)
    return _merge_bbox_candidates(img_body.size[0], img_body.size[1], body_bbox, scaled_cont)
def _tighten_pair_bbox(body_crop, cont_crop, pad_px):
    if body_crop.size != cont_crop.size:
        return None
    body_bbox = _meaningful_bbox_or_none(find_outer_bbox(body_crop.convert("RGB")), *body_crop.size)
    cont_bbox = _meaningful_bbox_or_none(find_outer_bbox(cont_crop.convert("RGB")), *cont_crop.size)
    tight_bbox = _merge_bbox_candidates(body_crop.size[0], body_crop.size[1], body_bbox, cont_bbox)
    if tight_bbox is None:
        return None
    x0, y0, x1, y1 = tight_bbox
    return clamp_bbox(x0 - pad_px, y0 - pad_px, x1 + pad_px, y1 + pad_px, body_crop.size[0], body_crop.size[1])
def _mix_size_from_crop(base_size, tight_size, outer_w_mm, outer_h_mm):
    bw, bh = base_size or (0, 0)
    tw, th = tight_size or (0, 0)
    if bw <= 0 or bh <= 0 or tw <= 0 or th <= 0:
        return float(outer_w_mm), float(outer_h_mm)
    scale = min(float(outer_w_mm) / float(bw), float(outer_h_mm) / float(bh))
    if scale <= 0:
        return float(outer_w_mm), float(outer_h_mm)
    return float(tw) * scale, float(th) * scale
def _compute_x_starts_for_w(w):
    W_limit = _single_print_content_w_mm() - float(SINGLE_MARGIN_R)
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
    H_limit = _single_print_content_h_max_mm()
    boundary = 464.0
    gap = 10.0
    y = float(QR_BAND + SINGLE_QR_IMAGE_GAP_MM)
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
        W_limit = _single_print_content_w_mm() - float(SINGLE_MARGIN_R)
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
def _right_align_single_x_starts(x_starts, w):
    starts = [float(x) for x in (x_starts or [])]
    if (not starts) or float(w) <= 0:
        return starts
    last_end = max(starts) + float(w)
    shift = _single_print_content_w_mm() - float(SINGLE_MARGIN_R) - float(last_end)
    if abs(shift) <= 1e-9:
        return starts
    return [float(x) + float(shift) for x in starts]
def build_single_placements_full_rows(best_base, rows_seg, align_right=True):
    w = float(best_base["w"])
    h = float(best_base["h"])
    s = int(best_base["k_cols"])
    rotated = (int(best_base["ori"]) == 1)
    x_starts = list(best_base.get("x_starts") or [])
    if align_right:
        x_starts = _right_align_single_x_starts(x_starts, w)
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
        H_seg = float(QR_BAND + SINGLE_QR_IMAGE_GAP_MM)
    seg_best = dict(best_base)
    seg_best["x_starts"] = list(x_starts)
    seg_best["rows_per_sheet"] = int(rows_seg)
    seg_best["items_per_sheet"] = int(rows_seg) * int(s)
    seg_best["H"] = float(H_seg)
    return placements, seg_best
def draw_label_in_band(page, text, x_mm, y_mm, w_mm):
    fontfile = pick_cjk_fontfile()
    txt = str(text or "").replace("\r", " ").replace("\n", " ").strip()
    if not txt:
        return
    xpt = mm_to_pt(float(x_mm) + 1.0)
    # 整拼标签固定单行，贴着第一条横向刀线上方
    ypt = mm_to_pt(float(y_mm) - 0.8)
    _draw_label_lines(page, [txt], xpt, ypt, LABEL_FONT_SIZE, fontfile)
def _draw_single_special_label(page, text, x_mm, y_mm):
    fontfile = pick_cjk_fontfile()
    txt = str(text or "").replace("\r", " ").replace("\n", " ").strip()
    if not txt:
        return
    xpt = mm_to_pt(float(x_mm))
    ypt = mm_to_pt(float(y_mm)) + float(LABEL_FONT_SIZE)
    _draw_label_lines(page, [txt], xpt, ypt, LABEL_FONT_SIZE, fontfile)
def _draw_single_segment_markers(page, seg, yoff_mm):
    placements = seg["placements"]
    label_text = seg["label_text"]
    qr_bytes = seg["qr_bytes"]
    qr_x0, qr_y0, qr_x1, qr_y1 = _single_qr_rect_mm(seg, yoff_mm)
    qr_rect = fitz.Rect(mm_to_pt(qr_x0), mm_to_pt(qr_y0), mm_to_pt(qr_x1), mm_to_pt(qr_y1))
    if label_text and placements:
        if _is_single_special_mode(seg):
            _draw_single_special_label(
                page,
                label_text,
                float(SINGLE_SPECIAL_LABEL_X_MM),
                float(yoff_mm) + float(SINGLE_SPECIAL_LABEL_Y_MM),
            )
        else:
            first = placements[0]
            draw_label_in_band(page, label_text, first["x"], yoff_mm + first["y"], first["w"])
    if qr_bytes:
        _draw_qr_on_white(page, qr_rect, qr_bytes)
def _merge_intervals(intervals, tol=1e-3):
    items = sorted((float(a), float(b)) for a, b in intervals if float(b) > float(a) + float(tol))
    if not items:
        return []
    merged = [list(items[0])]
    for a, b in items[1:]:
        if a <= merged[-1][1] + float(tol):
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return [(a, b) for a, b in merged]
def _placement_line_maps(placements, yoff_mm=0.0, tol=1e-3):
    items = list(placements or [])
    if not items:
        return {}, {}
    vertical = {}
    horizontal = {}
    for p in items:
        x0 = float(p["x"])
        y0 = float(yoff_mm) + float(p["y"])
        x1 = x0 + float(p["w"])
        y1 = y0 + float(p["h"])
        vertical.setdefault(round(x0, 6), []).append((y0, y1))
        vertical.setdefault(round(x1, 6), []).append((y0, y1))
        horizontal.setdefault(round(y0, 6), []).append((x0, x1))
        horizontal.setdefault(round(y1, 6), []).append((x0, x1))
    vertical = {
        float(x_key): _merge_intervals(intervals, tol=tol)
        for x_key, intervals in vertical.items()
    }
    horizontal = {
        float(y_key): _merge_intervals(intervals, tol=tol)
        for y_key, intervals in horizontal.items()
    }
    return vertical, horizontal
def _draw_line_maps(page, vertical, horizontal, color, width_pt, protrude_mm=0.0):
    if (not vertical) and (not horizontal):
        return
    protrude_mm = max(0.0, float(protrude_mm))
    shape = page.new_shape()
    for x, intervals in vertical.items():
        for y0, y1 in intervals:
            shape.draw_line(
                fitz.Point(mm_to_pt(x), mm_to_pt(y0 - protrude_mm)),
                fitz.Point(mm_to_pt(x), mm_to_pt(y1 + protrude_mm)),
            )
    for y, intervals in horizontal.items():
        for x0, x1 in intervals:
            shape.draw_line(
                fitz.Point(mm_to_pt(x0 - protrude_mm), mm_to_pt(y)),
                fitz.Point(mm_to_pt(x1 + protrude_mm), mm_to_pt(y)),
            )
    shape.finish(color=color, width=float(width_pt))
    shape.commit()
def _draw_line_map_extensions(page, vertical, horizontal, color, width_pt, protrude_mm=0.0):
    protrude_mm = max(0.0, float(protrude_mm))
    if protrude_mm <= 1e-9:
        return
    shape = page.new_shape()
    for x, intervals in vertical.items():
        for y0, y1 in intervals:
            shape.draw_line(
                fitz.Point(mm_to_pt(x), mm_to_pt(y0 - protrude_mm)),
                fitz.Point(mm_to_pt(x), mm_to_pt(y0)),
            )
            shape.draw_line(
                fitz.Point(mm_to_pt(x), mm_to_pt(y1)),
                fitz.Point(mm_to_pt(x), mm_to_pt(y1 + protrude_mm)),
            )
    for y, intervals in horizontal.items():
        for x0, x1 in intervals:
            shape.draw_line(
                fitz.Point(mm_to_pt(x0 - protrude_mm), mm_to_pt(y)),
                fitz.Point(mm_to_pt(x0), mm_to_pt(y)),
            )
            shape.draw_line(
                fitz.Point(mm_to_pt(x1), mm_to_pt(y)),
                fitz.Point(mm_to_pt(x1 + protrude_mm), mm_to_pt(y)),
            )
    shape.finish(color=color, width=float(width_pt))
    shape.commit()
def _single_cut_contour_line_maps(seg, yoff_mm):
    if bool(seg.get("single_is_double_cut", True)):
        return {}, {}
    return _placement_line_maps(seg.get("placements") or [], yoff_mm=yoff_mm)
def _draw_single_cut_contour_ticks(page, seg, yoff_mm, page_h_mm, exclude_x_range=None):
    vertical, horizontal = _single_cut_contour_line_maps(seg, yoff_mm)
    if (not vertical) and (not horizontal):
        return
    lo = hi = None
    if exclude_x_range is not None:
        lo, hi = exclude_x_range
        lo = float(lo)
        hi = float(hi)
    for x in sorted(vertical):
        if lo is not None and hi is not None and (lo + 1e-6) < float(x) < (hi - 1e-6):
            continue
        if 0.0 - 1e-6 <= float(x) <= _single_print_content_w_mm() + 1e-6:
            _draw_top_tick_at_y(page, float(x), yoff_mm, tick_len_mm=MARK_LEN, lw=1.0)
    for y in sorted(horizontal):
        if 0.0 - 1e-6 <= float(y) <= float(page_h_mm) + 1e-6:
            _draw_left_edge_tick(page, float(y), tick_len_mm=MARK_LEN, lw=1.0)
def _draw_single_cut_contour_grid(page, seg, yoff_mm):
    vertical, horizontal = _single_cut_contour_line_maps(seg, yoff_mm)
    if (not vertical) and (not horizontal):
        return
    _draw_line_maps(
        page,
        vertical,
        horizontal,
        color=(0.45, 0.83, 0.38),
        width_pt=SINGLE_CUT_CONTOUR_LINE_W_PT,
        protrude_mm=CONTOUR_LINE_EXT_MM,
    )
def _draw_single_segment_on_page_no_cuts(page, seg, img_bytes, yoff_mm, src_spec=None, src_doc_cache=None, is_contour=False):
    placements = seg["placements"]
    inner_margin_mm = float(seg.get("single_inner_margin_mm", INNER_MARGIN_MM))
    is_double_cut = bool(seg.get("single_is_double_cut", True))
    if is_contour and (not is_double_cut):
        return
    img_stream_cache = {False: img_bytes}
    edge_trim_mm = float(seg.get("single_shared_line_trim_mm", SINGLE_CUT_SHARED_LINE_TRIM_MM))
    for p in placements:
        x, y, w, h = p["x"], (yoff_mm + p["y"]), p["w"], p["h"]
        outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
        if DRAW_PART_OUTER_BOX:
            page.draw_rect(outer, color=PDF_CMYK_BLACK, width=0.5)
        inner = fitz.Rect(
            outer.x0 + mm_to_pt(inner_margin_mm),
            outer.y0 + mm_to_pt(inner_margin_mm),
            outer.x1 - mm_to_pt(inner_margin_mm),
            outer.y1 - mm_to_pt(inner_margin_mm),
        )
        target_rect = outer if (is_contour and (not is_double_cut)) else inner
        shown = False
        draw_src_spec = src_spec
        trim_left_mm = 0.0
        trim_top_mm = 0.0
        if is_contour and (not is_double_cut):
            has_left_neighbor = any(
                abs((float(q["x"]) + float(q["w"])) - float(p["x"])) <= 1e-3 and
                abs(float(q["y"]) - float(p["y"])) <= 1e-3
                for q in placements
            )
            has_top_neighbor = any(
                abs((float(q["y"]) + float(q["h"])) - float(p["y"])) <= 1e-3 and
                abs(float(q["x"]) - float(p["x"])) <= 1e-3
                for q in placements
            )
            if has_left_neighbor:
                trim_left_mm = edge_trim_mm
            if has_top_neighbor:
                trim_top_mm = edge_trim_mm
        if is_contour and (not is_double_cut):
            draw_src_spec = _trim_src_spec_clip_by_png_content(src_spec, img_bytes)
            draw_src_spec = _trim_src_spec_edges_mm(
                draw_src_spec,
                trim_left_mm=trim_left_mm,
                trim_top_mm=trim_top_mm,
            )
        if draw_src_spec:
            shown = _show_pdf_clip(
                page,
                target_rect,
                draw_src_spec,
                src_doc_cache or {},
                rotate_deg=(90 if p["rot"] else 0),
                keep_proportion=(not (is_contour and (not is_double_cut))),
            )
        if (not shown) and img_bytes:
            rot_flag = bool(p["rot"])
            if rot_flag not in img_stream_cache:
                img_to_use = img_bytes
                if is_contour and (not is_double_cut):
                    img_to_use = _crop_png_bytes_to_content(img_to_use)
                    img_to_use = _crop_png_bytes_by_src_trim(
                        img_to_use,
                        _trim_src_spec_clip_by_png_content(src_spec, img_bytes),
                        trim_left_mm=trim_left_mm,
                        trim_top_mm=trim_top_mm,
                    )
                img_stream_cache[rot_flag] = _rotate_png_bytes_if_needed(img_to_use, rot_flag)
            page.insert_image(
                target_rect,
                stream=img_stream_cache[rot_flag],
                keep_proportion=(not (is_contour and (not is_double_cut))),
            )
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
def _draw_single_print_page_corners(page, page_w_mm, page_h_mm):
    corner_len_pt = mm_to_pt(float(SINGLE_PRINT_CORNER_LEN_MM))
    page_w_pt = mm_to_pt(float(page_w_mm))
    page_h_pt = mm_to_pt(float(page_h_mm))
    lw = float(SINGLE_PRINT_CORNER_W_PT)
    page.draw_line(fitz.Point(0, 0), fitz.Point(corner_len_pt, 0), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(0, 0), fitz.Point(0, corner_len_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(page_w_pt - corner_len_pt, 0), fitz.Point(page_w_pt, 0), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(page_w_pt, 0), fitz.Point(page_w_pt, corner_len_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(0, page_h_pt - corner_len_pt), fitz.Point(0, page_h_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(0, page_h_pt), fitz.Point(corner_len_pt, page_h_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(page_w_pt - corner_len_pt, page_h_pt), fitz.Point(page_w_pt, page_h_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(page_w_pt, page_h_pt - corner_len_pt), fitz.Point(page_w_pt, page_h_pt), color=PDF_CMYK_BLACK, width=lw)

def _draw_single_print_seg_number(page, idx, y_mm, page_h_mm):
    box_w_mm = max(2.4, float(SINGLE_PRINT_SIDE_MARGIN_MM) - 0.8)
    box_h_mm = max(float(SINGLE_PRINT_SEG_MARK_EDGE_LEN_MM), 10.0)
    y0_mm = max(0.0, float(y_mm) - box_h_mm / 2.0)
    y1_mm = min(float(page_h_mm), float(y_mm) + box_h_mm / 2.0)
    x0_mm = 0.4 - float(SINGLE_PRINT_SEG_NUMBER_SHIFT_LEFT_MM)
    rect = fitz.Rect(
        mm_to_pt(x0_mm),
        mm_to_pt(y0_mm),
        mm_to_pt(x0_mm + box_w_mm),
        mm_to_pt(y1_mm),
    )
    page.insert_textbox(
        rect,
        str(int(idx)),
        fontsize=float(SINGLE_PRINT_SEG_MARK_FONT_PT),
        rotate=270,
        align=1,
        **_single_print_seg_text_kwargs(page)
    )

def _draw_single_print_seg_marker(page, page_w_mm, page_h_mm, y_mm, idx, draw_number):
    edge_len_mm = float(SINGLE_PRINT_SEG_MARK_EDGE_LEN_MM)
    arm_len_mm = float(SINGLE_PRINT_SIDE_MARGIN_MM)
    lw = float(SINGLE_PRINT_SEG_MARK_W_PT)
    y0_mm = max(0.0, float(y_mm) - edge_len_mm / 2.0)
    y1_mm = min(float(page_h_mm), float(y_mm) + edge_len_mm / 2.0)
    y_ctr_mm = float(y_mm)
    x_left_pt = 0.0
    x_right_pt = mm_to_pt(float(page_w_mm))
    y0_pt = mm_to_pt(y0_mm)
    y1_pt = mm_to_pt(y1_mm)
    y_ctr_pt = mm_to_pt(y_ctr_mm)
    arm_pt = mm_to_pt(arm_len_mm)
    page.draw_line(fitz.Point(x_left_pt, y0_pt), fitz.Point(x_left_pt, y1_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(x_left_pt, y_ctr_pt), fitz.Point(x_left_pt + arm_pt, y_ctr_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(x_right_pt, y0_pt), fitz.Point(x_right_pt, y1_pt), color=PDF_CMYK_BLACK, width=lw)
    page.draw_line(fitz.Point(x_right_pt - arm_pt, y_ctr_pt), fitz.Point(x_right_pt, y_ctr_pt), color=PDF_CMYK_BLACK, width=lw)
    if not draw_number:
        return
    _draw_single_print_seg_number(
        page,
        idx,
        min(float(page_h_mm), float(y_mm) + float(SINGLE_PRINT_SEG_NUMBER_SHIFT_DOWN_MM)),
        page_h_mm,
    )

def _draw_single_print_segment_marks(page, page_w_mm, page_h_mm):
    total_h = float(page_h_mm)
    step = float(SINGLE_PRINT_SEG_MARK_STEP_MM)
    if step <= 0:
        return
    seg_count = max(1, int(math.ceil(total_h / step)))
    for idx in range(1, seg_count + 1):
        y_mm = float((idx - 1) * step)
        if y_mm > total_h + 1e-6:
            continue
        if idx == 1:
            _draw_single_print_seg_number(page, 1, float(SINGLE_PRINT_CORNER_LEN_MM) + 3.0, total_h)
            continue
        _draw_single_print_seg_marker(page, page_w_mm, page_h_mm, y_mm, idx, True)

def _render_one_single_sheet_content_doc(sheet, is_contour, fixed_h_mm, draw_page_border=True):
    W = _single_print_content_w_mm()
    H = float(fixed_h_mm)
    doc = fitz.open()
    page = doc.new_page(width=mm_to_pt(W), height=mm_to_pt(H))
    if draw_page_border:
        page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=PDF_CMYK_BLACK, width=1.0)
    yoff = 0.0
    src_doc_cache = {}
    try:
        for idx, seg in enumerate(sheet["segments"]):
            img_bytes = seg["img_cont"] if is_contour else seg["img_body"]
            src_spec = seg.get("src_cont") if is_contour else seg.get("src_body")
            use_single_special_ticks = bool((not is_contour) and _is_single_special_mode(seg))
            if idx > 0:
                if not is_contour:
                    _draw_horizontal_cutline(page, yoff, lw=1.0)
            _draw_single_segment_on_page_no_cuts(
                page, seg, img_bytes, yoff_mm=yoff, src_spec=src_spec, src_doc_cache=src_doc_cache, is_contour=is_contour
            )
            if is_contour:
                _draw_single_cut_contour_grid(page, seg, yoff_mm=yoff)
            if not is_contour:
                if use_single_special_ticks:
                    _draw_single_special_edge_ticks(page, seg, yoff_mm=yoff, page_h_mm=H, top_only=True)
                else:
                    _draw_segment_cut_ticks(page, seg, yoff_mm=yoff, page_h_mm=H, draw_left_ticks=False)
                _draw_single_segment_markers(page, seg, yoff_mm=yoff)
            if (not is_contour) and seg.get("qr_bytes"):
                qr_x0, _qr_y0, qr_x1, _qr_y1 = _single_qr_rect_mm(seg, yoff)
                if use_single_special_ticks:
                    _draw_single_special_edge_ticks(
                        page,
                        seg,
                        yoff_mm=yoff,
                        page_h_mm=H,
                        exclude_x_range=(qr_x0, qr_x1),
                        top_only=True,
                    )
                else:
                    _redraw_single_top_ticks(page, seg, yoff_mm=yoff, exclude_x_range=(qr_x0, qr_x1))
            yoff += float(seg["best"]["H"])
    finally:
        _close_doc_cache(src_doc_cache)
    return doc

def _render_one_single_sheet_doc(sheet, is_contour, fixed_h_mm=SINGLE_H_MAX):
    content_h_mm = float(fixed_h_mm)
    outer_w_mm = float(SINGLE_PRINT_PAGE_W_MM)
    outer_h_mm = min(
        float(SINGLE_H_MAX),
        content_h_mm + float(SINGLE_PRINT_TOP_MARGIN_MM) + float(SINGLE_PRINT_BOTTOM_MARGIN_MM),
    )
    content_doc = _render_one_single_sheet_content_doc(
        sheet,
        bool(is_contour),
        content_h_mm,
        draw_page_border=False,
    )
    work_doc = fitz.open()
    try:
        page = work_doc.new_page(width=mm_to_pt(outer_w_mm), height=mm_to_pt(outer_h_mm))
        target_rect = fitz.Rect(
            mm_to_pt(float(SINGLE_PRINT_SIDE_MARGIN_MM)),
            mm_to_pt(float(SINGLE_PRINT_TOP_MARGIN_MM)),
            mm_to_pt(float(SINGLE_PRINT_SIDE_MARGIN_MM) + _single_print_content_w_mm()),
            mm_to_pt(float(SINGLE_PRINT_TOP_MARGIN_MM) + content_h_mm),
        )
        page.show_pdf_page(target_rect, content_doc, 0, keep_proportion=False)
        if not is_contour:
            _draw_single_outer_left_cut_ticks(page, sheet, outer_h_mm)
            _draw_single_print_page_corners(page, outer_w_mm, outer_h_mm)
            _draw_single_print_segment_marks(page, outer_w_mm, outer_h_mm)
        final_doc = fitz.open()
        final_page = final_doc.new_page(width=mm_to_pt(outer_w_mm), height=mm_to_pt(outer_h_mm))
        final_page.show_pdf_page(
            fitz.Rect(0, 0, mm_to_pt(outer_w_mm), mm_to_pt(outer_h_mm)),
            work_doc,
            0,
            rotate=180,
            keep_proportion=False,
        )
        return final_doc
    finally:
        try:
            content_doc.close()
        except Exception:
            pass
        try:
            work_doc.close()
        except Exception:
            pass
def _pick_cmyk_icc_profile_path():
    color_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "spool", "drivers", "color")
    candidates = [
        "CoatedFOGRA39.icc",
        "CoatedFOGRA27.icm",
        "ISOcoated_v2_300_eci.icc",
        "PSOcoated_v3.icc",
        "GRACoL2006_Coated1v2.icc",
        "USWebCoatedSWOP.icc",
        "JapanColor2001Coated.icc",
    ]
    for name in candidates:
        path = os.path.join(color_dir, name)
        if os.path.exists(path):
            return path
    return None
def _pdf_catalog_xref(doc):
    try:
        return int(doc.pdf_catalog())
    except Exception:
        pass
    try:
        trailer = doc.pdf_trailer()
        m = re.search(r"/Root\s+(\d+)\s+0\s+R", str(trailer))
        if m:
            return int(m.group(1))
    except Exception:
        pass
    try:
        _kind, root_ref = doc.xref_get_key(-1, "Root")
        m = re.search(r"(\d+)\s+0\s+R", str(root_ref))
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return 0
def _new_pdf_xref(doc):
    for name in ("get_new_xref", "new_xref", "_newXref"):
        fn = getattr(doc, name, None)
        if callable(fn):
            try:
                return int(fn())
            except Exception:
                pass
    raise RuntimeError("no xref allocator")
def _best_effort_attach_cmyk_output_intent(pdf_path):
    icc_path = _pick_cmyk_icc_profile_path()
    if not icc_path or (not os.path.exists(pdf_path)):
        return
    try:
        with open(icc_path, "rb") as f:
            icc_bytes = f.read()
        if not icc_bytes:
            return
        doc = fitz.open(pdf_path)
    except Exception:
        return
    tmp_path = None
    try:
        catalog_xref = _pdf_catalog_xref(doc)
        if catalog_xref <= 0:
            return
        icc_xref = _new_pdf_xref(doc)
        doc.update_object(
            icc_xref,
            "<< /N 4 /Alternate /DeviceCMYK /Length %d >>" % len(icc_bytes),
        )
        doc.update_stream(icc_xref, icc_bytes)
        oi_xref = _new_pdf_xref(doc)
        profile_name = os.path.splitext(os.path.basename(icc_path))[0]
        doc.update_object(
            oi_xref,
            "<< /Type /OutputIntent /S /GTS_PDFX /OutputConditionIdentifier (%s) "
            "/Info (%s) /RegistryName (https://www.color.org) /DestOutputProfile %d 0 R >>"
            % (profile_name, profile_name, icc_xref),
        )
        if hasattr(doc, "xref_set_key"):
            doc.xref_set_key(catalog_xref, "OutputIntents", "[%d 0 R]" % oi_xref)
        else:
            catalog_obj = doc.xref_object(catalog_xref, compressed=False)
            if "/OutputIntents" in catalog_obj:
                catalog_obj = re.sub(r"/OutputIntents\s*\[[^\]]*\]", "/OutputIntents [%d 0 R]" % oi_xref, catalog_obj, flags=re.S)
            else:
                catalog_obj = catalog_obj.rstrip()
                if catalog_obj.endswith(">>"):
                    catalog_obj = catalog_obj[:-2] + "\n/OutputIntents [%d 0 R]\n>>" % oi_xref
            doc.update_object(catalog_xref, catalog_obj)
        try:
            doc.saveIncr()
        except Exception:
            tmp_path = pdf_path + ".intent.tmp.pdf"
            doc.save(tmp_path, garbage=4, deflate=True, incremental=False)
    except Exception:
        pass
    finally:
        try:
            doc.close()
        except Exception:
            pass
    if tmp_path and os.path.exists(tmp_path):
        try:
            os.replace(tmp_path, pdf_path)
        except Exception:
            try:
                shutil.copyfile(tmp_path, pdf_path)
            finally:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
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
        saved_path = out_path
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
        saved_path = alt_pdf
    _best_effort_attach_cmyk_output_intent(saved_path)
    return saved_path
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
    page_body = 2 * pair_index
    page_cont = 2 * pair_index + 1
    pad_px = mm_to_px(CROP_PAD_MM, RENDER_DPI)
    tight_pad_px = max(1, mm_to_px(0.5, RENDER_DPI))
    body_full = render_page_to_pil(path, page_index=page_body, dpi=RENDER_DPI, doc=doc)
    cont_full = render_page_to_pil(path, page_index=page_cont, dpi=RENDER_DPI, doc=doc)
    body_page_rect = doc.load_page(page_body).rect
    cont_page_rect = doc.load_page(page_cont).rect
    shared_bbox = _detect_shared_pair_bbox(doc, page_body, page_cont, body_full, cont_full)
    cont_bbox = _scale_bbox_between_sizes(shared_bbox, body_full.size, cont_full.size)
    body_base_bbox = _bbox_with_pad_or_full(shared_bbox, body_full.size, pad_px=pad_px)
    cont_base_bbox = _bbox_with_pad_or_full(cont_bbox, cont_full.size, pad_px=pad_px)
    body_crop = body_full.crop(body_base_bbox).convert("RGBA")
    cont_crop = cont_full.crop(cont_base_bbox).convert("RGBA")
    base_crop_size = body_crop.size
    tight_bbox = _tighten_pair_bbox(body_crop, cont_crop, pad_px=tight_pad_px)
    if tight_bbox is not None:
        body_crop = body_crop.crop(tight_bbox).convert("RGBA")
        cont_crop = cont_crop.crop(tight_bbox).convert("RGBA")
        bx0, by0, bx1, by1 = body_base_bbox
        cx0, cy0, cx1, cy1 = cont_base_bbox
        tx0, ty0, tx1, ty1 = tight_bbox
        body_final_bbox = clamp_bbox(bx0 + tx0, by0 + ty0, bx0 + tx1, by0 + ty1, body_full.size[0], body_full.size[1])
        cont_final_bbox = clamp_bbox(cx0 + tx0, cy0 + ty0, cx0 + tx1, cy0 + ty1, cont_full.size[0], cont_full.size[1])
    else:
        body_final_bbox = body_base_bbox
        cont_final_bbox = cont_base_bbox
    return _pil_to_png_bytes(body_crop), _pil_to_png_bytes(cont_crop), {
        "base_size": tuple(base_crop_size),
        "tight_size": tuple(body_crop.size),
        "src_body": _build_src_spec_from_px_bbox(path, page_body, body_final_bbox, body_page_rect, body_full.size),
        "src_cont": _build_src_spec_from_px_bbox(path, page_cont, cont_final_bbox, cont_page_rect, cont_full.size),
    }
def _save_dual_docs(doc1, doc2, out_path1, out_path2):
    return safe_save(doc1, out_path1), safe_save(doc2, out_path2)
def _append_skip(skip_list, log_cb, type_name_raw, reason):
    skip_list.append((type_name_raw, reason))
    _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, reason))
def _finish_stage_progress(progress_cb, total):
    if progress_cb:
        progress_cb(total, total, "完成 %d/%d" % (total, total))
def _iter_open_pdf_jobs(archived_paths, progress_cb=None, on_error=None):
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
    for pi in range(pair_count):
        img_body, img_cont, crop_info = _build_pair_images(path, doc, pi)
        tid = "%s@P%d" % (type_name_raw, pi + 1)
        yield tid, img_body, img_cont, crop_info
def _run_single_stage(archived_paths, progress_cb=None, log_cb=None):
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
            single_rule = get_single_layout_rule_from_filename(path)
            outer_ext_mm = float(single_rule.get("outer_ext_mm", OUTER_EXT))
            outer_w0 = float(A + 2.0 * outer_ext_mm)
            outer_h0 = float(B + 2.0 * outer_ext_mm)
            mix_outer_w0 = float(A + 2.0 * OUTER_EXT)
            mix_outer_h0 = float(B + 2.0 * OUTER_EXT)
            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)
            qr_bytes = make_qr_png_bytes(qr_text)
            if N < MIX_THRESHOLD:
                skip_single.append((type_name_raw, "N<%d_skip_no_layout" % MIX_THRESHOLD))
                _log(log_cb, "⏭️ N<%d 不拼版跳过：%s  N=%d" % (MIX_THRESHOLD, type_name_raw, N))
                for tid, img_body, img_cont, crop_info in _iter_pair_payloads(path, d, type_name_raw, pair_count):
                    mix_w0, mix_h0 = _mix_size_from_crop(
                        (crop_info or {}).get("base_size"),
                        (crop_info or {}).get("tight_size"),
                        mix_outer_w0,
                        mix_outer_h0,
                    )
                    Mi, Ni, rot90 = _mix_choose_orient(mix_w0, mix_h0)
                    mix_pool.append({
                        "tid": tid,
                        "label_text": label_text,
                        "qr_bytes": qr_bytes,
                        "img_body": img_body,
                        "img_cont": img_cont,
                        "src_body": (crop_info or {}).get("src_body"),
                        "src_cont": (crop_info or {}).get("src_cont"),
                        "rem": int(N),
                        "W": float(Mi),
                        "H": float(Ni),
                        "rot90": bool(rot90),
                    })
                    ok_mix_pool.append((tid, "mix_pool"))
                    _log(log_cb, "🧩 MIX_POOL: %s N=%d Mi=%.1f Ni=%.1f rot90=%s src=%.1fx%.1f tight=%.1fx%.1f"
                         % (tid, int(N), float(Mi), float(Ni), str(bool(rot90)),
                            float(outer_w0), float(outer_h0), float(mix_w0), float(mix_h0)))
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
            for tid, img_body, img_cont, crop_info in _iter_pair_payloads(path, d, type_name_raw, pair_count):
                remain_rows = r_total
                seg_cnt = 0
                single_align_right = not (bool(SINGLE_SINGLE_LINE_SPECIAL_MODE) and (not bool(single_rule.get("is_double_cut", True))))
                while remain_rows > 0:
                    seg_cnt += 1
                    rows_seg = min(max_rows, remain_rows)
                    placements, seg_best = build_single_placements_full_rows(best_base, rows_seg, align_right=single_align_right)
                    global_single_segments.append({
                        "type_key": tid,
                        "tid": tid,
                        "best": seg_best,
                        "placements": placements,
                        "qr_bytes": qr_bytes,
                        "label_text": label_text,
                        "img_body": img_body,
                        "img_cont": img_cont,
                        "src_body": (crop_info or {}).get("src_body"),
                        "src_cont": (crop_info or {}).get("src_cont"),
                        "single_is_double_cut": bool(single_rule.get("is_double_cut", True)),
                        "single_inner_margin_mm": float(single_rule.get("inner_margin_mm", INNER_MARGIN_MM)),
                        "single_outer_ext_mm": float(outer_ext_mm),
                        "single_output_name": build_single_output_name(path),
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
        packed = _pack_segments_ffd_to_fixed_sheets(global_single_segments, max_h_mm=_single_print_content_h_max_mm(), gap_mm=0.0)
        single_name_counts = {}
        single_counter = 0
        for sh in packed:
            single_counter += 1
            name = _build_single_sheet_pdf_name(sh, single_counter, single_name_counts)
            out_path1 = os.path.join(DEST_DIR1, name)
            out_path2 = os.path.join(DEST_DIR2, name)
            page_h_mm = float(sh.get("used_h", 0.0))
            if page_h_mm <= 0:
                page_h_mm = _single_print_content_h_max_mm()
            if any(_is_single_special_mode(seg) for seg in (sh.get("segments") or [])):
                page_h_mm += float(SINGLE_SPECIAL_BOTTOM_CLEAR_MM)
            doc1 = _render_one_single_sheet_doc(sh, is_contour=False, fixed_h_mm=page_h_mm)
            p1 = safe_save(doc1, out_path1)
            doc2 = _render_one_single_sheet_doc(sh, is_contour=True, fixed_h_mm=page_h_mm)
            p2 = safe_save(doc2, out_path2)
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
def _mix_marker_key(t):
    return str((t or {}).get("label_text") or (t or {}).get("tid") or "")
def _mix_effective_image_band(row_y, marker_band):
    band = max(0.0, float(marker_band))
    if band <= 1e-9:
        return 0.0
    return max(_mix_min_marker_band_mm(), band) if abs(float(row_y) - float(MARGIN_TOP_MM)) <= 1e-6 else band
def _advance_mix_y(y, row_used_h):
    y = float(y)
    y2 = y + float(row_used_h)
    boundary = 464.0 * (math.floor((y + 1e-9) / 464.0) + 1.0)
    return y2 + 10.0 if y < boundary and y2 > boundary + 1e-9 else float(y2)
def _mix_overlap_1d(a0, a1, b0, b1, eps=1e-6):
    return min(float(a1), float(b1)) > max(float(a0), float(b0)) + float(eps)
def _mix_lane_bounds(block):
    if float(block["x1"]) <= float(SPLIT_X) + 1e-6:
        return float(PAGE_MARGIN_L), float(SPLIT_X)
    return float(SPLIT_X + SPLIT_GAP_W), float(PAGE_W - PAGE_MARGIN_R)
def _mix_band_floor(y0):
    floor = float(MARGIN_TOP_MM)
    boundary = 464.0
    while boundary + 10.0 <= float(y0) + 1e-6:
        floor = boundary + 10.0
        boundary += 464.0
    return floor
def _mix_build_block_groups(page):
    groups = [{"block": b, "placements": []} for b in (page.get("blocks") or [])]
    if not groups:
        return groups
    for p in (page.get("placements") or []):
        best = None
        best_area = None
        for g in groups:
            b = g["block"]
            if p.get("type") is not b.get("type"):
                continue
            if float(p["x"]) + 1e-6 < float(b["x0"]) or float(p["x"]) + float(p["w"]) > float(b["x1"]) + 1e-6:
                continue
            if float(p["y"]) + 1e-6 < float(b["y0"]) or float(p["y"]) + float(p["h"]) > float(b["y1"]) + 1e-6:
                continue
            area = (float(b["x1"]) - float(b["x0"])) * (float(b["y1"]) - float(b["y0"]))
            if best is None or area < best_area:
                best = g
                best_area = area
        if best is not None:
            best["placements"].append(p)
    return groups
def _mix_move_group(group, dx, dy):
    b = group["block"]
    dx = float(dx)
    dy = float(dy)
    b["x0"] = float(b["x0"]) + dx
    b["x1"] = float(b["x1"]) + dx
    b["y0"] = float(b["y0"]) + dy
    b["y1"] = float(b["y1"]) + dy
    for p in group["placements"]:
        p["x"] = float(p["x"]) + dx
        p["y"] = float(p["y"]) + dy
def _mix_refresh_used_h(page):
    max_end = float(MARGIN_TOP_MM)
    for b in (page.get("blocks") or []):
        max_end = max(max_end, float(b["y1"]))
    for p in (page.get("placements") or []):
        max_end = max(max_end, float(p["y"]) + float(p["h"]))
    page["used_h"] = float(max_end)
def _mix_compact_page(page, gap_x, gap_y):
    groups = _mix_build_block_groups(page)
    if not groups:
        return
    gap_x = float(gap_x)
    gap_y = float(gap_y)
    for _ in range(4):
        moved = False
        for g in sorted(groups, key=lambda it: (float(it["block"]["x0"]), float(it["block"]["y0"]))):
            b = g["block"]
            lane_left, lane_right = _mix_lane_bounds(b)
            width = float(b["x1"]) - float(b["x0"])
            target_x0 = lane_left
            for og in groups:
                if og is g:
                    continue
                ob = og["block"]
                if not _mix_overlap_1d(b["y0"], b["y1"], ob["y0"], ob["y1"]):
                    continue
                if float(ob["x1"]) <= float(b["x0"]) + 1e-6:
                    target_x0 = max(target_x0, float(ob["x1"]) + gap_x)
            target_x0 = min(target_x0, lane_right - width)
            if target_x0 < float(b["x0"]) - 1e-6:
                _mix_move_group(g, target_x0 - float(b["x0"]), 0.0)
                moved = True
        for g in sorted(groups, key=lambda it: (float(it["block"]["y0"]), float(it["block"]["x0"]))):
            b = g["block"]
            target_y0 = _mix_band_floor(b["y0"])
            for og in groups:
                if og is g:
                    continue
                ob = og["block"]
                if not _mix_overlap_1d(b["x0"], b["x1"], ob["x0"], ob["x1"]):
                    continue
                if float(ob["y1"]) <= float(b["y0"]) + 1e-6:
                    target_y0 = max(target_y0, float(ob["y1"]) + gap_y)
            if target_y0 < float(b["y0"]) - 1e-6:
                _mix_move_group(g, 0.0, target_y0 - float(b["y0"]))
                moved = True
        if not moved:
            break
    _mix_refresh_used_h(page)
def _collect_mix_base_heights(types_list, max_w, marker_keys=None, row_y=None):
    marker_keys = set(marker_keys or [])
    hs = set()
    for t in types_list:
        if int(t.get("rem", 0)) <= 0:
            continue
        marker_band = _mix_marker_band_for_width(
            t, min(float(max_w), float(t.get("W", 0.0))), marker_keys
        )
        image_band = _mix_effective_image_band(row_y, marker_band)
        for wi, hi, _rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)):
            if float(wi) <= float(max_w) + 1e-9:
                hs.add(float(hi) + image_band)
    return sorted(hs)
def pack_mix_by_height_rule(mix_types, page_max_h=PAGE_H_MAX):
    content_max_h = max(1.0, float(page_max_h) - float(MARGIN_BOTTOM_MM))
    X_LEFT0 = float(PAGE_MARGIN_L)
    X_LEFT1 = float(SPLIT_X)
    X_RIGHT0 = float(SPLIT_X + SPLIT_GAP_W)
    X_RIGHT1 = float(PAGE_W - PAGE_MARGIN_R)
    gap_x = float(MIX_GAP_X_MM)
    gap_y = float(CELL_GAP_MM)
    types_list = [dict(t, rem=int(t.get("rem", t.get("N", 0)) or 0)) for t in mix_types]
    for t in types_list:
        fit_lane = any(float(wi) <= float(LEFT_BLOCK_W) + 1e-9
                       for wi, _hi, _rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)))
        if int(t["rem"]) > 0 and (not fit_lane):
            print("⚠️ MIX_SKIP_TOO_WIDE_FOR_LANE:", t.get("tid"), "minW=%.2f > %.2f" % (min(float(t["W"]), float(t["H"])), float(LEFT_BLOCK_W)))
            t["rem"] = 0
    def _pick_orient_and_layout(t, base_h, rem_w, marker_keys, row_y):
        best = None
        best_rank = None
        rem = int(t.get("rem", 0))
        marker_key = _mix_marker_key(t)
        marker_keys_set = set(marker_keys or [])
        for Wi, Hi, rot90 in _iter_mix_orients(t["W"], t["H"], t.get("rot90", False)):
            if Wi > rem_w + 1e-9:
                continue
            max_cols = int(math.floor((rem_w + gap_x) / (Wi + gap_x))) if Wi > 0 else 0
            if max_cols <= 0:
                continue
            marker_band = 0.0 if marker_key in marker_keys_set else _mix_min_marker_band_mm()
            cols = 1
            stack_cap = 1
            place_cnt = 1
            block_w = float(Wi)
            for _ in range(4):
                image_band = _mix_effective_image_band(row_y, marker_band)
                avail_h = float(base_h) - float(image_band)
                if avail_h <= 1e-9 or Hi > avail_h + 1e-9:
                    stack_cap = 0
                    break
                stack_cap = int(math.floor((avail_h + gap_y) / (Hi + gap_y))) if Hi > 0 else 1
                if stack_cap <= 0:
                    stack_cap = 1
                need_cols = int(math.ceil(float(rem) / float(stack_cap))) if rem > 0 else 1
                cols = min(max_cols, need_cols)
                if cols <= 0:
                    cols = 1
                place_cnt = min(rem, cols * stack_cap)
                block_w = float(cols) * Wi + float(max(0, cols - 1)) * gap_x
                new_marker_band = _mix_marker_band_for_width(t, block_w, marker_keys_set)
                if abs(float(new_marker_band) - float(marker_band)) <= 0.1:
                    marker_band = float(new_marker_band)
                    break
                marker_band = float(new_marker_band)
            image_band = _mix_effective_image_band(row_y, marker_band)
            avail_h = float(base_h) - float(image_band)
            if stack_cap <= 0 or avail_h <= 1e-9 or Hi > avail_h + 1e-9:
                continue
            need_cols = int(math.ceil(float(rem) / float(stack_cap))) if rem > 0 else 1
            cols = min(max_cols, need_cols)
            if cols <= 0:
                cols = 1
            place_cnt = min(rem, cols * stack_cap)
            block_w = float(cols) * Wi + float(max(0, cols - 1)) * gap_x
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
                    "marker_band": float(marker_band),
                    "image_band": float(image_band),
                    "marker_key": marker_key,
                }
                best_rank = rank
        return best
    def _build_one_region_row(work_list, base_h, y, x_start, x_end, page_marker_keys):
        x = float(x_start)
        row_blocks = []
        row_added = []
        row_marker_keys = set(page_marker_keys or [])
        new_marker_keys = set()
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
                plan = _pick_orient_and_layout(t, base_h, rem_w, row_marker_keys, y)
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
            marker_band = float(cand_plan.get("marker_band", QR_BAND))
            image_band = float(cand_plan.get("image_band", marker_band))
            marker_key = str(cand_plan.get("marker_key") or "")
            x0 = float(x)
            block_w = float(cand_plan["block_w"])
            x1 = x0 + block_w
            block_min_x = None
            block_max_x = None
            block_max_y = None
            for c in range(cols):
                col_x = x0 + float(c) * (Wi + gap_x)
                for s in range(stack_cap):
                    if int(cand["rem"]) <= 0:
                        break
                    px = col_x
                    py = float(y) + image_band + float(s) * (Hi + gap_y)
                    row_added.append(_mix_make_placement(
                        cand, px, py, Wi, Hi, cand_plan["rot90"], marker_band, image_band
                    ))
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
                rowBlocks = _mix_make_block(
                    cand, block_min_x, block_max_x, y, block_max_y, marker_band, image_band
                )
                row_blocks.append(rowBlocks)
                if marker_band > 1e-9 and marker_key:
                    row_marker_keys.add(marker_key)
                    new_marker_keys.add(marker_key)
            x = x1 + gap_x
        if not row_added:
            return [], [], 0.0, 0.0, set()
        row_max_end = 0.0
        for p in row_added:
            row_max_end = max(row_max_end, float(p["y"]) + float(p["h"]))
        row_used_h = max(0.0, row_max_end - float(y))
        return row_blocks, row_added, float(row_used_h), float(row_max_end), new_marker_keys
    def _find_best_region_row(work_list, region, page_marker_keys):
        region_w = float(region["x1"] - region["x0"])
        if region_w <= 1e-6 or float(region["y"]) >= float(content_max_h) - 1e-6:
            return None
        hs = _collect_mix_base_heights(work_list, region_w, marker_keys=page_marker_keys, row_y=region["y"])
        cand_hs = hs
        best = None
        best_score = None
        for bh in cand_hs:
            sim_list = [dict(t) for t in work_list]
            row_blocks, row_added, row_used_h, row_max_end, new_marker_keys = _build_one_region_row(
                sim_list, bh, region["y"], region["x0"], region["x1"], set(page_marker_keys or [])
            )
            if not row_added:
                continue
            y_after = _advance_mix_y(region["y"], row_used_h)
            if max(float(row_max_end), float(y_after)) > float(content_max_h) + 1e-9:
                continue
            max_x1 = max(float(b["x1"]) for b in row_blocks) if row_blocks else float(region["x0"])
            leftover = max(0.0, float(region["x1"]) - max_x1)
            used_area = sum(float(p["w"]) * float(p["h"]) for p in row_added)
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
                    "new_marker_keys": set(new_marker_keys),
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
        page = {"blocks": [], "placements": [], "used_h": 0.0, "marker_keys": set()}
        page_marker_keys = set()
        regions = [
            {"name": "L", "x0": X_LEFT0, "x1": X_LEFT1, "y": float(MARGIN_TOP_MM)},
            {"name": "R", "x0": X_RIGHT0, "x1": X_RIGHT1, "y": float(MARGIN_TOP_MM)},
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
                best = _find_best_region_row(types_list, region, page_marker_keys)
                if best is None:
                    continue
                other_y = float(regions[1 - ridx]["y"])
                projected_page_h = max(float(page["used_h"]), float(best["row_max_end"]), float(best["y_after"]), other_y)
                balance_gap = abs(float(best["y_after"]) - other_y)
                score = (
                    float(region["y"]),
                    float(projected_page_h),
                    float(balance_gap),
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
            row_blocks2, row_added2, row_used_h2, row_max_end2, new_marker_keys2 = _build_one_region_row(
                types_list, best["base_h"], region["y"], region["x0"], region["x1"], page_marker_keys
            )
            if not row_added2:
                break
            page["blocks"].extend(row_blocks2)
            page["placements"].extend(row_added2)
            page_marker_keys.update(new_marker_keys2)
            page["marker_keys"] = set(page_marker_keys)
            region["y"] = _advance_mix_y(region["y"], row_used_h2)
            page["used_h"] = max(
                page["used_h"],
                float(row_max_end2),
                float(region["y"]),
                float(regions[1 - ridx]["y"]),
            )
            if min(float(r["y"]) for r in regions) >= float(content_max_h) - 1e-6:
                break
        if not page["placements"]:
            types_list = [t for t in types_list if int(t["rem"]) > 0]
            if not types_list:
                break
            types_list.sort(key=lambda a: (-float(a.get("W", 0.0)), -float(a.get("H", 0.0))))
            print("⚠️ MIX_SKIP_UNPLACEABLE:", types_list[0].get("tid"))
            types_list[0]["rem"] = 0
            continue
        _mix_refresh_used_h(page)
        _mix_compact_page(page, gap_x, gap_y)
        _mix_hole_fill(page, types_list, page_max_h, gap_x=gap_x, gap_y=gap_y)
        _mix_compact_page(page, gap_x, gap_y)
        _mix_refresh_used_h(page)
        page["used_h"] = float(max(float(MARGIN_TOP_MM), min(float(content_max_h), float(page["used_h"]))))
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
def _mix_hole_fill(page, types_list, page_max_h, gap_x, gap_y):
    H = float(page.get("used_h", 0.0))
    if H <= 1e-6:
        return
    page_marker_keys = set(page.get("marker_keys") or [])
    usable_x0 = float(PAGE_MARGIN_L)
    usable_x1 = float(PAGE_W - PAGE_MARGIN_R)
    usable_y0 = float(MARGIN_TOP_MM)
    usable_y1 = max(float(usable_y0) + 1.0, float(page_max_h) - float(MARGIN_BOTTOM_MM))
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
        ow = int(round(w + gap_x))
        oh = int(round(h + gap_y))
        bp.cut_out(ox, oy, ow, oh)
    for p in page.get("placements", []):
        x = float(p["x"]) - usable_x0
        y = float(p["y"]) - usable_y0
        w = float(p["w"])
        h = float(p["h"])
        ox = int(round(x))
        oy = int(round(y))
        ow = int(round(w + gap_x))
        oh = int(round(h + gap_y))
        bp.cut_out(ox, oy, ow, oh)
    cand_types = [t for t in types_list if int(t.get("rem", 0)) > 0]
    cand_types.sort(key=lambda t: -(float(t["W"]) * (float(t["H"]) + float(QR_BAND))))
    added = 0
    guard = 0
    while True:
        guard += 1
        if guard > 200000:
            break
        best_choice = None
        for t in cand_types:
            if int(t["rem"]) <= 0:
                continue
            w0 = float(t["W"])
            h0 = float(t["H"])
            marker_key = _mix_marker_key(t)
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
                marker_band = _mix_marker_band_for_width(t, wi, page_marker_keys)
                image_band = float(marker_band)
                total_h = float(image_band) + float(hi)
                rw = int(round(wi))
                rh = int(round(total_h))
                pos = bp.find_bottom_left(rw + int(round(gap_x)), rh + int(round(gap_y)))
                if pos is None:
                    continue
                rank = (
                    int(pos["y"]),
                    int(pos["x"]),
                    -int(round(wi * total_h)),
                    abs(int(round(wi)) - int(round(total_h))),
                )
                if best_pick is None or rank < best_pick[0]:
                    best_pick = (rank, pos, wi, hi, total_h, rot90, marker_band, image_band, marker_key)
            if best_pick is None:
                continue
            if best_choice is None or best_pick[0] < best_choice[0]:
                best_choice = best_pick + (t,)
        if best_choice is None:
            break
        _rank, pos, wi, hi, total_h, rot90, marker_band, image_band, marker_key, t = best_choice
        bp.cut_out(pos["x"], pos["y"], int(round(wi + gap_x)), int(round(total_h + gap_y)))
        block_x0 = float(pos["x"]) + usable_x0
        block_y0 = float(pos["y"]) + usable_y0
        img_y = float(block_y0) + float(image_band)
        page["blocks"].append(_mix_make_block(
            t, block_x0, block_x0 + wi, block_y0, img_y + hi, marker_band, image_band
        ))
        page["placements"].append(_mix_make_placement(
            t, block_x0, img_y, wi, hi, rot90, marker_band, image_band
        ))
        if marker_band > 1e-9 and marker_key:
            page_marker_keys.add(marker_key)
            page["marker_keys"] = set(page_marker_keys)
        t["rem"] -= 1
        added += 1
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
                "marker_band": float(it.get("marker_band", QR_BAND)),
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
                "marker_band": 0.0,
            })
    if not blocks:
        return
    min_y0 = min(float(b["y0"]) for b in blocks)
    top_blocks = []
    for b in blocks:
        if abs(float(b["y0"]) - min_y0) <= eps:
            top_blocks.append(b)
    xs = set()
    top_x_min = None
    top_x_max = None
    if top_blocks:
        raw_xs = []
        for b in top_blocks:
            raw_xs.append(float(b["x0"]))
            raw_xs.append(float(b["x1"]))
        if raw_xs:
            top_x_min = min(raw_xs)
            top_x_max = max(raw_xs)
    for b in top_blocks:
        for x_raw in (float(b["x0"]), float(b["x1"])):
            x_draw = float(x_raw)
            if top_x_min is not None and abs(x_draw - top_x_min) <= 1e-6:
                x_draw = _snap_edge_mm(x_draw, 0.0, float(PAGE_W), EDGE_SNAP_MM)
            elif top_x_max is not None and abs(x_draw - top_x_max) <= 1e-6:
                x_draw = _snap_edge_mm(x_draw, 0.0, float(PAGE_W), EDGE_SNAP_MM)
            xs.add(round(x_draw, 3))
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
        marker_band = float(b.get("marker_band", QR_BAND))
        top_cut_y = float(b["y0"]) + marker_band + 0.8
        bottom_cut_y = float(b["y1"])
        if marker_band > 1e-6 and 0.0 - 1e-6 <= top_cut_y <= float(page_h_mm) + 1e-6:
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
    page.draw_rect(fitz.Rect(0, 0, mm_to_pt(W), mm_to_pt(H)), color=PDF_CMYK_BLACK, width=1.0)
    placements = page_obj.get("placements", [])
    img_stream_cache = {}
    src_doc_cache = {}
    try:
        for p in placements:
            t = p["type"]
            img_bytes = t["img_cont"] if is_contour else t["img_body"]
            src_spec = t.get("src_cont") if is_contour else t.get("src_body")
            x = float(p["x"])
            y = float(p["y"])
            w = float(p["w"])
            h = float(p["h"])
            rot90 = bool(p.get("rot90", False))
            outer = fitz.Rect(mm_to_pt(x), mm_to_pt(y), mm_to_pt(x + w), mm_to_pt(y + h))
            if DRAW_PART_OUTER_BOX:
                page.draw_rect(outer, color=PDF_CMYK_BLACK, width=0.5)
            shown = False
            if src_spec:
                shown = _show_pdf_clip(page, outer, src_spec, src_doc_cache, rotate_deg=(90 if rot90 else 0))
            if (not shown) and img_bytes:
                cache_key = (id(t), bool(is_contour), rot90)
                if cache_key not in img_stream_cache:
                    img_stream_cache[cache_key] = _rotate_png_bytes_if_needed(img_bytes, rot90)
                page.insert_image(outer, stream=img_stream_cache[cache_key], keep_proportion=True)
        if is_contour:
            vertical, horizontal = _placement_line_maps(placements, yoff_mm=0.0)
            _draw_line_map_extensions(
                page,
                vertical,
                horizontal,
                color=(0.45, 0.83, 0.38),
                width_pt=SINGLE_CUT_CONTOUR_LINE_W_PT,
                protrude_mm=CONTOUR_LINE_EXT_MM,
            )
        else:
            _mix_draw_edge_ticks(page, page_obj.get("blocks", []), page_h_mm=H)
            qr_draws = []
            for b in page_obj.get("blocks", []):
                if not bool(b.get("needs_marker", False)):
                    continue
                t = b["type"]
                x0 = float(b["x0"])
                x1 = float(b["x1"])
                y0 = float(b["y0"])
                qr_bytes = t.get("qr_bytes", None)
                if qr_bytes:
                    qr_rect = fitz.Rect(
                        mm_to_pt(x1 - QR_W), mm_to_pt(y0 + MIX_QR_TOP_GAP_MM),
                        mm_to_pt(x1), mm_to_pt(y0 + MIX_QR_TOP_GAP_MM + QR_H)
                    )
                    qr_draws.append((qr_rect, qr_bytes))
                label_text = t.get("label_text", "")
                if label_text:
                    label_band_h = max(float(QR_BAND), float(b.get("marker_band", QR_BAND)))
                    _draw_label_top_left(
                        page,
                        label_text,
                        x0,
                        y0,
                        max_w_mm=_mix_label_available_w_mm(x1 - x0),
                        band_h_mm=label_band_h,
                    )
            for qr_rect, qr_bytes in qr_draws:
                _draw_qr_on_white(page, qr_rect, qr_bytes)
        return doc
    finally:
        _close_doc_cache(src_doc_cache)
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
            p1 = safe_save(doc1, out_path1)
            doc2 = render_mix_page(pg, is_contour=True)
            p2 = safe_save(doc2, out_path2)
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
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2, SINGLE_SINGLE_LINE_SPECIAL_MODE
    if cfg:
        DEST_DIR = cfg.get("DEST_DIR", DEST_DIR)
        IN_PDF_ARCHIVE_DIR = cfg.get("TEST_DIR", IN_PDF_ARCHIVE_DIR)
        DEST_DIR1 = cfg.get("DEST_DIR1", DEST_DIR1)
        DEST_DIR2 = cfg.get("DEST_DIR2", DEST_DIR2)
        SINGLE_SINGLE_LINE_SPECIAL_MODE = bool(
            cfg.get("SINGLE_SINGLE_LINE_SPECIAL_MODE", SINGLE_SINGLE_LINE_SPECIAL_MODE)
        )
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
