# -*- coding: utf-8 -*-
import os

# =========================
# 依赖检测（保持原逻辑）
# =========================
try:
    import cv2  # noqa
    CV2_OK = True
except Exception:
    CV2_OK = False

try:
    import qrcode  # noqa
    QR_OK = True
except Exception:
    QR_OK = False

try:
    from PIL import ImageDraw  # noqa
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

OUT_PDF_P1 = os.path.join(DEST_DIR1, "over_test_p1.pdf")
OUT_PDF_P2 = os.path.join(DEST_DIR2, "over_test_p2.pdf")


def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2, OUT_PDF_P1, OUT_PDF_P2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)
    OUT_PDF_P1 = os.path.join(DEST_DIR1, "over_test_p1.pdf")
    OUT_PDF_P2 = os.path.join(DEST_DIR2, "over_test_p2.pdf")


# =========================
# 工艺参数
# =========================
OUTER_EXT = 2.0
INNER_GAP = 3.0

GAP = 6.0
LABEL_BAND = 6.0

INNER_MARGIN_MM = 5.0
MARK_LEN = 5.0

# 图形块尺寸（一块的最大宽高）
BLOCK_W = 320.0
BLOCK_H = 464.0
BLOCK_GAP = 10.0   # 块与块之间的间隔(mm)

# 纸张尺寸限制
SINGLE_W_MIN = 600.0
SINGLE_W_MAX = 600.0
SINGLE_H_MIN = 1
SINGLE_H_MAX = 1500

FULL_W_MIN = 600
FULL_W_MAX = 600
FULL_H_MAX = 1500

# N < FULL_THRESHOLD 时直接进入全拼
FULL_THRESHOLD = 10

QR_BAND = 10.0
QR_W = 10.0
QR_H = 10.0

LABEL_FONT_SIZE = 10

RENDER_DPI = 600
CROP_PAD_MM = 1.2
TEXT_MASK_PAD_MM = 1.2

DRAW_PART_OUTER_BOX = False

# 全拼 band
FULL_TOP_PAD = 0.0
FULL_LABEL_BAND = 10.0
