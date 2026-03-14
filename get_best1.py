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
OUTER_EXT = 2.0                              # 外接框扩张毫米数，裁剪时在自动识别的 bbox 基础上额外扩张，避免裁得太紧。过大可能导致多余空白，过小可能裁得太紧。2mm 大约等于 5-6 像素（在 600 DPI 下），是一个较好的平衡点。
CELL_GAP_MM = 10.0                           # 图与图之间的间距
MARGIN_LR_MM = 5.0                           # 左右边距
MARGIN_TOP_MM = 10.0                         # 上边距
MARGIN_BOTTOM_MM = 5.0                       # 下边距
INNER_MARGIN_MM = 0.0                        # 图块内部边距，裁剪时在每个图块内额外保留的空白，避免内容过于靠近边缘。过大可能导致图块尺寸不够用，过小可能导致内容太靠近边缘。0mm 表示不额外保留空白，直接裁到识别的 bbox 边界，是一个较为激进的设置。
MARK_LEN = 5.0                               # 刀线/刻度长度，单位毫米。过长可能导致刀线过于显眼，过短可能不够明显。5mm 是一个较好的平衡点。
PAGE_W = 600.0                               # 页面宽度
PAGE_H_MAX = 1500.0                          # 页面最大高度，超过则分成多页
PAGE_MARGIN_L = MARGIN_LR_MM                 
PAGE_MARGIN_R = MARGIN_LR_MM
SPLIT_X = 320.0                              # 320mm 处的分界线，超过则跨页；不足则混拼在同一页
SPLIT_GAP_W = 10.0                           # 分界线保留的空白宽度(320~330)，跨页时中间留空；混拼时不留空
QR_BAND = MARGIN_TOP_MM
QR_W = 10.0                                  # 二维码宽度
QR_H = 10.0                                  # 二维码高度
LABEL_FONT_SIZE = 10                         # 标签字体大小(单位pt)，会自动按块宽度截断文字，避免超出边界
RENDER_DPI = 600                             # PDF渲染分辨率，单位 DPI。过低可能导致结构识别不准确，过高则处理慢且占内存。600 是一个较好的平衡点。                
CROP_PAD_MM = 1.2                            # 裁剪时在 bbox 四周额外扩张的毫米数，避免裁得太紧。过大可能导致多余空白，过小可能裁得太紧。1.2mm 大约等于 3-4 像素（在 600 DPI 下），是一个较好的平衡点。
TEXT_MASK_PAD_MM = 1.2                       # 文本区域扩张的毫米数，用于在渲染图上涂白文本区域，避免被误识别为图形。过大可能导致过多图形被遮盖，过小可能遮盖不全。1.2mm 大约等于 3-4 像素（在 600 DPI 下），是一个较好的平衡点。
SAFE_PAD_MM = 3.0                            # 安全边距毫米数，用于在最终输出的图块中保留额外的空白，避免内容过于靠近边缘。过大可能导致图块尺寸不够用，过小可能导致内容太靠近边缘。3mm 大约等于 7-8 像素（在 600 DPI 下），是一个较好的平衡点。
DRAW_PART_OUTER_BOX = False
MIX_ENABLE = True
MIX_THRESHOLD = 10  # N < 10 -> mix；N >= 10 -> skip

LEFT_BLOCK_W = float(SPLIT_X - PAGE_MARGIN_L)  # 315
TOTAL_USABLE_W = float((PAGE_W - PAGE_MARGIN_R) - PAGE_MARGIN_L)  # 590

EDGE_SNAP_MM = 10.0                          # 刀线吸边毫米数，允许刀线自动吸附到边界，避免出现残缺的刀线。过大可能导致刀线位置不合理，过小可能无法有效吸边。10mm 大约等于 24 像素（在 600 DPI 下），是一个较好的平衡点。
def set_runtime_paths(dest_dir, archive_dir, out_dir1, out_dir2):
    """
    功能：
    动态设置运行时输入/输出目录，通常由外部 run.py 调用。
    """
    global DEST_DIR, IN_PDF_ARCHIVE_DIR, DEST_DIR1, DEST_DIR2
    DEST_DIR = str(dest_dir)
    IN_PDF_ARCHIVE_DIR = str(archive_dir)
    DEST_DIR1 = str(out_dir1)
    DEST_DIR2 = str(out_dir2)

def ensure_dir(p):                            # 确保目录存在的工具函数，避免重复代码
    if p and (not os.path.isdir(p)):          # 避免p为空或已存在但不是目录的情况
        os.makedirs(p)                        # 创建目录

def mm_to_pt(mm):
    return mm * 72.0 / 25.4                   # mm 转 pt 的公式，72.0/25.4 是一个常数约等于 2.83465

def mm_to_px(mm, dpi):                        # 当前版本中未直接使用，保留以便后续需要时调用
    return int(round(mm * dpi / 25.4))        # mm 转 px 的公式，dpi/25.4 是一个常数约等于 23.62205（在 600 DPI 下）

def _log(log_cb, s):                          # 统一的日志输出函数，接受一个 log_cb 回调函数和一个字符串 s
    if log_cb:                                # 仅当 log_cb 不为 None 时才调用，避免传入非函数类型导致错误                           
        log_cb(s)                             # 回调输出日志，允许外部 run.py 捕获并统一管理日志
    else:
        print(s)                              # 直接打印日志，适用于没有传入 log_cb 的情况

def clamp_bbox(x0, y0, x1, y1, img_w, img_h): # 将 bbox 坐标限制在图像边界内，避免裁剪越界。过大可能导致越界错误，过小可能导致内容被裁掉。这个函数会确保 bbox 至少有 1 像素的宽高，并且完全在图像范围内。
    x0 = max(0, min(img_w - 1, int(round(x0))))
    y0 = max(0, min(img_h - 1, int(round(y0))))
    x1 = max(x0 + 1, min(img_w, int(round(x1))))
    y1 = max(y0 + 1, min(img_h, int(round(y1))))
    return x0, y0, x1, y1

def union_bbox(b1, b2):                       #  
    """
    功能：
    合并两个 bbox,返回它们的并集。
    任意一个为 None 时，返回另一个。
    """
    if b1 is None:                            # 如果 b1 是 None，直接返回 b2，无需计算
        return b2
    if b2 is None:                            # 如果 b2 是 None，直接返回 b1，无需计算
        return b1
    return (
        min(b1[0], b2[0]),
        min(b1[1], b2[1]),
        max(b1[2], b2[2]),
        max(b1[3], b2[3]),
    )

def rect_union(rects):                        
    rects = [r for r in rects if r is not None] # 过滤掉 None 的矩形，避免计算错误
    if not rects:                               # 如果没有有效的矩形，返回 None，表示没有区域
        return None
    x0 = min(r.x0 for r in rects)               # 计算所有矩形的最小 x0，作为外接矩形的左边界
    y0 = min(r.y0 for r in rects)               # 计算所有矩形的最小 y0，作为外接矩形的上边界
    x1 = max(r.x1 for r in rects)               # 计算所有矩形的最大 x1，作为外接矩形的右边界
    y1 = max(r.y1 for r in rects)               # 计算所有矩形的最大 y1，作为外接矩形的下边界
    return fitz.Rect(x0, y0, x1, y1)            # 返回一个新的 fitz.Rect，表示所有输入矩形的外接矩形

def archive_input_pdf_to_dir(src_pdf_path, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)       # 确保目标目录存在，避免后续复制文件时出错
    base = os.path.basename(src_pdf_path)     # 获取源 PDF 的文件名，保留扩展名，用于构造目标路径
    dst_path = os.path.join(dst_dir, base)    # 构造目标路径，初始假设与源文件同名，后续会根据情况调整

    try:
        if os.path.abspath(src_pdf_path).lower() == os.path.abspath(dst_path).lower():   # 比较源路径和目标路径的绝对路径（忽略大小写），如果相同则直接返回目标路径，无需复制
            return dst_path
    except Exception:
        pass

    if os.path.exists(dst_path):              # 如果目标路径已存在，可能是之前复制过的文件，需要检查是否与源文件相同
        try:
            if os.path.getsize(dst_path) == os.path.getsize(src_pdf_path):              
             # 如果目标文件和源文件大小相同，基本可以认为是同一个文件，直接复用目标路径，无需复制
                return dst_path
        except Exception:
            pass
        name, ext = os.path.splitext(base)    # 如果目标文件存在但与源文件不同，则需要重命名目标文件，避免覆盖。这里使用当前时间戳来生成一个唯一的文件名，格式为 "原文件名_年月日_时分秒.扩展名"，确保不会与现有文件冲突。
        ts = time.strftime("%Y%m%d_%H%M%S")   # 生成当前时间的字符串，格式为 "年月日_时分秒"，例如 "20240601_153045"，用于构造新的文件名
        # 构造新的目标路径，格式为 "目标目录/原文件名_时间戳.扩展名"，例如 "D:/test_data/test/图纸_20240601_153045.pdf"，确保不会覆盖现有文件
        dst_path = os.path.join(dst_dir, "%s_%s%s" % (name, ts, ext))                    

    shutil.copy2(src_pdf_path, dst_path)      # 复制源 PDF 到目标路径，使用 copy2 保留文件的元数据（如修改时间等），确保归档后的文件与原文件尽可能一致
    return dst_path

def _draw_left_edge_tick(page, y_mm, tick_len_mm=MARK_LEN, lw=1.0):    # 在纸张最左侧画一条水平短刀线（从左边缘往右伸）
    ypt = mm_to_pt(float(y_mm))                                        # 将 y_mm 转换为 PDF 坐标系中的 pt 单位，fitz 使用的坐标系是以左上角为原点，单位为 pt（1/72 英寸），因此需要将毫米转换为 pt 来正确定位刀线位置
    tick = mm_to_pt(float(tick_len_mm))                                # 将 tick_len_mm 转换为 pt 单位，表示刀线的长度，确保在 PDF 中正确显示
    # 在 PDF 页面上绘制一条线段，起点是 (0, ypt)，终点是 (tick, ypt)，颜色为黑色，线宽为 lw。这样就形成了一条从左边缘向右伸展的水平短刀线，位置由 y_mm 决定，长度由 tick_len_mm 决定。
    page.draw_line(fitz.Point(0, ypt), fitz.Point(tick, ypt), color=(0, 0, 0), width=lw)

def _draw_top_edge_tick(page, x_mm, tick_len_mm=MARK_LEN, lw=1.0):     # 在纸张最顶部画一条竖直短刀线（从上边缘往下伸）
    xpt = mm_to_pt(float(x_mm))                                        # 将 x_mm 转换为 PDF 坐标系中的 pt 单位，fitz 使用的坐标系是以左上角为原点，单位为 pt（1/72 英寸），因此需要将毫米转换为 pt 来正确定位刀线位置
    tick = mm_to_pt(float(tick_len_mm))                                # 将 tick_len_mm 转换为 pt 单位，表示刀线的长度，确保在 PDF 中正确显示
    # 在 PDF 页面上绘制一条线段，起点是 (xpt, 0)，终点是 (xpt, tick)，颜色为黑色，线宽为 lw。这样就形成了一条从上边缘往下伸展的竖直短刀线，位置由 x_mm 决定，长度由 tick_len_mm 决定。
    page.draw_line(fitz.Point(xpt, 0), fitz.Point(xpt, tick), color=(0, 0, 0), width=lw)

def parse_A_B_N_from_filename(pdf_path):                               # 从文件名中解析尺寸 A×B 和数量 N
    base = os.path.splitext(os.path.basename(pdf_path))[0]             # 获取文件名（不带路径和扩展名），例如 "图纸_A×B^N"，用于后续解析
    parts = base.split("^")                                            # 将文件名按 "^" 分割成多个部分，得到一个列表，例如 ["图纸_A×B", "N"]，用于后续查找尺寸和数量信息

    size_part = None                                                   # 初始化 size_part 和 n_part 变量，用于存储解析到的尺寸和数量信息
    n_part = None                                                      # 如果文件名格式规范，且分割后至少有 8 个部分，则直接使用第 7 和第 8 部分作为尺寸和数量信息，避免复杂的正则匹配，提高解析效率
    # 如果文件名格式不规范，或者分割后部分不足，则使用正则表达式在所有部分中查找符合 A×B 格式的尺寸信息和数字格式的数量信息，确保能够解析出正确的尺寸和数量
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
    # 使用正则表达式从 size_part 中提取 A 和 B 的数值，要求格式为 A×B，其中 A 和 B 可以是整数或小数，且可以有任意数量的空格和不同的乘号符号（*、x、X、×、✕）。如果不符合格式，则抛出异常提示文件名格式错误。
    m = re.search(r"(\d+(\.\d+)?)\s*[\*xX×✕]\s*(\d+(\.\d+)?)", str(size_part))
    if not m:
        raise ValueError("尺寸段不是AxB格式：%s" % str(size_part))

    A = float(m.group(1))                                              # 从正则表达式的匹配结果中提取 A 和 B 的数值，转换为浮点数类型，便于后续计算和比较
    B = float(m.group(3))                                              # A 对应正则表达式中的第一个捕获组，B 对应第三个捕获组，中间的乘号和空格不包含在捕获组中，因此不会影响数值的提取

    m2 = re.search(r"(\d+)", str(n_part))                              # 使用正则表达式从 n_part 中提取数量 N 的数值，要求格式为纯数字，如果不符合格式，则抛出异常提示文件名格式错误
    if not m2:
        raise ValueError("数量段不是数字：%s" % str(n_part))
    N = int(m2.group(1))

    if A <= 0 or B <= 0 or N <= 0:                                     # 对解析到的 A、B、N 进行基本的合理性检查，要求它们都必须是正数，如果不满足条件，则抛出异常提示解析结果非法，帮助用户检查文件名格式是否正确
        raise ValueError("解析到的 A/B/N 非法：A=%s B=%s N=%s (%s)" % (A, B, N, base))

    return A, B, N

def extract_label_text(pdf_path):                                      # 从文件名中提取标签文字，优先使用前两段拼接，避免过长的标签影响显示；如果没有足够的段，则退化为前 40 个字符，确保标签不会过长。
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    ps = base.split("^")
    if len(ps) >= 2:
        return "^".join(ps[:2])
    return base[:40]

def extract_qr_text_from_filename(pdf_path):                           # 从文件名中提取二维码文本，优先使用第 10 段（如果存在且非空），其次尝试匹配 SJ 开头的模式，最后退化为前两段拼接或前 80 个字符，确保二维码文本具有一定的唯一性和识别性。
    base = os.path.splitext(os.path.basename(pdf_path))[0]             # 获取文件名（不带路径和扩展名），例如 "图纸_A×B^N^QR"，用于后续解析
    ps = base.split("^")                                               # 将文件名按 "^" 分割成多个部分，得到一个列表，例如 ["图纸_A×B", "N", "QR"]，用于后续查找二维码文本信息
    if len(ps) >= 10 and ps[9]:                                        # 如果分割后至少有 10 个部分，且第 10 个部分非空，则直接使用第 10 个部分作为二维码文本，避免复杂的正则匹配，提高解析效率
        return str(ps[9]).strip()                                      # 返回第 10 个部分的字符串形式，并去除首尾空白，作为二维码文本

    m = re.search(r"(SJ[0-9A-Za-z]+)", base)                           # 如果没有直接指定二维码文本，则使用正则表达式在整个文件名中查找符合 SJ 开头的模式，要求后面跟随至少一个数字或字母，如果找到则使用该匹配结果作为二维码文本，确保二维码具有一定的唯一性和识别性
    if m:                                                              # 如果正则表达式匹配成功，返回第一个捕获组的内容，即 SJ 开头的字符串，作为二维码文本
        return m.group(1)

    if len(ps) >= 2:                                                   # 如果没有找到 SJ 模式，则退化为使用前两段拼接的方式来生成二维码文本，确保二维码文本不会过长且具有一定的识别性；如果前两段不足，则退化为使用前 80 个字符，避免二维码文本过长导致扫描困难。
        return "^".join(ps[:2])

    return base[:80]

def make_qr_png_bytes(qr_text, box_px=240):                            # 生成二维码的 PNG 图片字节，使用 qrcode 库生成二维码图像，并将其调整为指定的像素大小；如果 qrcode 库不可用，则生成一个占位图像，显示 "QR" 字样，确保函数始终返回一个有效的 PNG 图片字节。
    if QR_OK:                                                          # 如果 qrcode 库可用，使用 qrcode 库生成二维码图像。首先创建一个 QRCode 对象，设置版本为 None（自动调整大小），错误纠正级别为 M，盒子大小为 8 像素，边距为 0。然后将 qr_text 添加到 QRCode 对象中，并调用 make(fit=True) 来生成二维码图像。接着将二维码图像转换为 RGB 模式，并调整大小为 box_px × box_px 像素。最后将图像保存到一个 BytesIO 对象中，以 PNG 格式保存，并返回 PNG 图片的字节内容。
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=0
        )
        qr.add_data(qr_text)                                           # 将 qr_text 添加到 QRCode 对象中，准备生成二维码图像
        qr.make(fit=True)                                              # 调用 make(fit=True) 来生成二维码图像，fit=True 表示自动调整二维码的版本和大小以适应输入的数据，确保二维码能够正确编码 qr_text
        # 将生成的二维码图像转换为 RGB 模式，并调整大小为 box_px × box_px 像素，以满足输出要求。然后将图像保存到一个 BytesIO 对象中，以 PNG 格式保存，并返回 PNG 图片的字节内容，确保函数输出一个有效的二维码图片字节。
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img = img.resize((box_px, box_px))                             # 调整二维码图像的大小为 box_px × box_px 像素，确保输出的二维码图片符合指定的尺寸要求
        bio = BytesIO()                                                # 创建一个 BytesIO 对象，用于在内存中保存 PNG 图片数据，避免使用临时文件，确保函数的效率和简洁性
        img.save(bio, format="PNG", optimize=True)                     # 将调整大小后的二维码图像保存到 BytesIO 对象中，以 PNG 格式保存，并启用 optimize 选项来优化 PNG 图片的大小，确保输出的二维码图片既符合尺寸要求又尽可能小
        return bio.getvalue()                                          # 返回 PNG 图片的字节内容，供后续使用，例如嵌入到 PDF 中，确保函数输出一个有效的二维码图片字节

    img = Image.new("RGB", (box_px, box_px), (255, 255, 255))          # 如果 qrcode 库不可用，创建一个新的 RGB 图像，大小为 box_px × box_px 像素，背景颜色为白色，作为占位图像，确保函数始终返回一个有效的 PNG 图片字节
    if PIL_DRAW_OK:                                                    # 如果 PIL 的 ImageDraw 模块可用，在占位图像上绘制一个黑色的矩形边框，并在中心位置绘制 "QR" 字样，提示这是一个二维码占位图像，确保占位图像具有一定的识别性和提示作用；如果 ImageDraw 模块不可用，则保持纯白色背景的占位图像。
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, box_px - 1, box_px - 1], outline=(0, 0, 0), width=3)
        d.text((box_px * 0.28, box_px * 0.40), "QR", fill=(0, 0, 0))
    bio = BytesIO()                                                    # 将占位图像保存到一个 BytesIO 对象中，以 PNG 格式保存，并启用 optimize 选项来优化 PNG 图片的大小，确保输出的占位图像既具有提示作用又尽可能小
    img.save(bio, format="PNG", optimize=True)                         # 将占位图像保存到 BytesIO 对象中，以 PNG 格式保存，并启用 optimize 选项来优化 PNG 图片的大小，确保输出的占位图像既具有提示作用又尽可能小
    return bio.getvalue()                                              # 返回 PNG 图片的字节内容，供后续使用，例如嵌入到 PDF 中，确保函数输出一个有效的占位图像字节

def pick_cjk_fontfile(): # 选择一个可用的 CJK 字体文件，用于在 PDF 上绘制标签文字。函数会尝试几个常见的字体路径，返回第一个存在的路径；如果没有找到任何可用的字体，则返回 None，调用方需要做好兼容处理，例如使用默认字体。
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]
    for p in candidates: # 遍历候选字体路径列表，检查每个路径是否存在，如果存在则返回该路径，确保函数能够找到一个可用的 CJK 字体文件；如果没有找到任何存在的路径，则返回 None，调用方需要做好兼容处理，例如使用默认字体。
        if os.path.exists(p):
            return p
    return None
# 根据字体大小和最大宽度，估算文本的宽度，并在必要时进行截断，确保绘制在 PDF 上的标签文字不会超出指定的最大宽度。
# 函数会根据一个经验公式来估算文本的宽度，如果估算值超过最大宽度，则计算需要保留多少字符，并在末尾添加省略号 "..." 来表示被截断的部分；
# 如果估算值不超过最大宽度，则直接返回原文本。
def trim_text_to_width(text, fontsize, max_width_pt):                  
    est = 0.55 * fontsize * len(text)                                  # 使用一个经验公式来估算文本的宽度，假设每个字符的平均宽度约为 0.55 倍的字体大小（单位 pt），然后乘以文本的字符数，得到文本的总宽度估算值（单位 pt）。这个公式是一个近似值，实际宽度可能会有所偏差，但对于大多数常见字体和文本长度来说是一个合理的估算。
    if est <= max_width_pt:                                            # 如果估算的文本宽度不超过最大宽度，则直接返回原文本，无需截断，确保标签文字能够完整显示在 PDF 上。
        return text
    keep = max(0, int(max_width_pt / (0.55 * fontsize)) - 3)           # 如果估算的文本宽度超过最大宽度，则需要进行截断。首先计算在保留省略号 "..." 的情况下，最多可以保留多少个字符。这个计算是通过将最大宽度除以每个字符的平均宽度（0.55 * fontsize）来得到总共可以容纳多少个字符，然后再减去 3 个字符的空间来给省略号留出位置。最后使用 max(0, ...) 来确保保留的字符数不会为负数，避免出现错误。
    if keep <= 0:                                                      # 如果计算得到的保留字符数小于或等于 0，说明即使只保留省略号 "..." 也无法满足最大宽度的要求，此时直接返回 "..."，表示文本被完全截断，确保标签文字不会超出指定的最大宽度。
        return "..."
    return text[:keep] + "..."
# 在 PDF 页面上指定的位置绘制标签文字，使用 pick_cjk_fontfile() 选择一个可用的 CJK 字体文件，并使用 trim_text_to_width() 
# 来确保文本不会超出指定的最大宽度。函数会将文本绘制在 (x_left_mm, y_top_mm) 的位置，字体大小为 LABEL_FONT_SIZE，颜色为黑色；
# 如果没有找到可用的 CJK 字体，则使用默认字体 "helv" 来绘制文本，确保标签文字能够正确显示在 PDF 上。
def _draw_label_top_left(page, text, x_left_mm, y_top_mm, max_w_mm):   
    fontfile = pick_cjk_fontfile()                                     # 选择一个可用的 CJK 字体文件，如果没有找到则返回 None，调用方需要做好兼容处理，例如使用默认字体。
    fontsize = LABEL_FONT_SIZE                                         # 设置字体大小为 LABEL_FONT_SIZE，单位为 pt，用于在 PDF 上绘制标签文字，确保标签文字具有适当的大小和可读性。
    max_w_pt = mm_to_pt(max(5.0, float(max_w_mm)))                     # 将最大宽度 max_w_mm 转换为 PDF 坐标系中的 pt 单位，确保在 PDF 上正确限制标签文字的宽度。这里使用 max(5.0, ...) 来确保最大宽度至少为 5mm，避免过小的宽度导致标签文字无法显示。
    txt = trim_text_to_width(text, fontsize, max_w_pt)                 # 使用 trim_text_to_width() 来确保文本不会超出指定的最大宽度，函数会根据字体大小和最大宽度来估算文本的宽度，并在必要时进行截断，添加省略号 "..." 来表示被截断的部分，确保标签文字能够正确显示在 PDF 上且不会超出边界。
    xpt = mm_to_pt(float(x_left_mm) + 1.0)                             # 将 x_left_mm 转换为 PDF 坐标系中的 pt 单位，并在基础上加上 1.0 mm 的偏移，确保标签文字不会紧贴边缘，具有一定的内边距，提升视觉效果。
    ypt = mm_to_pt(float(y_top_mm) + 7.5)                              # 将 y_top_mm 转换为 PDF 坐标系中的 pt 单位，并在基础上加上 7.5 mm 的偏移，确保标签文字不会紧贴边缘，具有一定的内边距，提升视觉效果。这里的 7.5 mm 是一个经验值，适用于大多数字体大小和标签位置，可以根据实际情况进行调整。
    if fontfile:                                                       # 如果找到了可用的 CJK 字体文件，则使用该字体文件来绘制文本，确保标签文字能够正确显示在 PDF 上；如果没有找到可用的 CJK 字体文件，则使用默认字体 "helv" 来绘制文本，确保标签文字能够正确显示在 PDF 上。
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontfile=fontfile, color=(0, 0, 0))
    else:
        page.insert_text(fitz.Point(xpt, ypt), txt, fontsize=fontsize, fontname="helv", color=(0, 0, 0))
# PDF 结构信息提取，包括文本框、图像和绘图对象的边界框，用于后续只提取实际图形区域，避免把说明文字等也裁进去。
def _is_probably_page_frame(rect_obj, page_rect):
    if rect_obj is None:                                               # 如果传入的矩形对象为 None，直接返回 False，表示无法判断该矩形是否可能是页面边框，确保函数的健壮性。
        return False
    tol = 2.0
    if rect_obj.width <= 0 or rect_obj.height <= 0:                    # 如果传入的矩形对象的宽度或高度小于或等于 0，直接返回 False，表示该矩形不可能是页面边框，因为页面边框应该具有正的宽度和高度，确保函数的健壮性。
        return False
    # 判断传入的矩形对象是否在宽度和高度上都接近页面矩形的大小，使用 0.985 的比例作为一个经验阈值，确保函数能够识别出那些几乎覆盖整个页面的矩形，这些矩形很可能是页面边框。
    full_like = (rect_obj.width >= 0.985 * page_rect.width and rect_obj.height >= 0.985 * page_rect.height)
    # 判断传入的矩形对象是否在四条边上都接近页面矩形的边界，使用一个容差值 tol 来允许一定的偏差，确保函数能够识别出那些虽然不完全覆盖整个页面但在边界上非常接近的矩形，这些矩形也很可能是页面边框。 
    touches_all_edges = (
        abs(rect_obj.x0 - page_rect.x0) <= tol and
        abs(rect_obj.y0 - page_rect.y0) <= tol and
        abs(rect_obj.x1 - page_rect.x1) <= tol and
        abs(rect_obj.y1 - page_rect.y1) <= tol
    )
    return full_like and touches_all_edges
# 从 PDF 页面中提取结构信息，包括页面矩形、文本框、图像边界框和绘图对象边界框。函数会加载指定页码的页面，获取页面矩形，
# 并遍历页面上的文本块和图像块来提取它们的边界框，同时过滤掉那些可能是页面边框的矩形。对于绘图对象，也会提取它们的边界
# 框，并过滤掉可能是页面边框的矩形。最后返回一个包含页面矩形、文本框列表、图像边界框和绘图对象边界框的字典，供后续使用，
# 例如在渲染图像时遮盖文本区域，或者在检测图形时只关注实际的图像区域。
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
# 将 PDF 页面渲染为 PIL 图像，使用 fitz 库加载 PDF 并渲染指定页码的页面为图像。函数会根据指定的 DPI 来计算缩放比例，
# 并将渲染结果转换为 PIL 图像返回。
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
# 将 PDF 坐标系中的矩形边界框转换为像素坐标系中的边界框，函数会根据页面矩形和图像的宽高来计算缩放比例，
# 并将 PDF 坐标系中的矩形边界框转换为像素坐标系中的边界框，确保在图像上正确定位 PDF 中的文本或图像区域。
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
# 在 PIL 图像上遮盖 PDF 中的文本区域，函数会根据 PDF 坐标系中的文本框列表和页面矩形来计算文本区域在图像上的位置，
# 并使用白色矩形遮盖这些区域，确保在后续的图像处理或图形检测中不会受到文本区域的干扰。
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
# 从二值化的掩码图像中提取边界框，函数会根据掩码中非零像素的位置来计算边界框的坐标，并进行一些基本的过滤，
# 例如过滤掉过小的区域或过大的区域，确保提取到的边界框具有一定的合理性和代表性。
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
# 判断一个边界框是否可能是页面边框，函数会检查边界框是否接近页面的四条边，并且是否在宽度和高度上都接近页面的大小，
# 确保能够识别出那些几乎覆盖整个页面或者在边界上非常接近的矩形，这些矩形很可能是页面边框。
def _reject_border_bbox(x, y, w, h, W, H):
    edge_pad = 8
    touches_edge = (x <= edge_pad or y <= edge_pad or (x + w) >= (W - edge_pad) or (y + h) >= (H - edge_pad))
    huge = (w >= 0.985 * W and h >= 0.985 * H)
    return touches_edge and huge
# 从二值化的掩码图像中提取边界框，函数会使用 OpenCV 的 findContours 方法来找到掩码中的轮廓，并根据轮廓计算边界框。
# 函数还会进行一些过滤，例如过滤掉过小的区域或过大的区域，以及可能是页面边框的区域,确保提取到的边界框具有一定的合理
# 性和代表性。最后函数会根据面积来选择合适的边界框，或者将多个边界框进行合并，返回一个最终的边界框。
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
# 从 PIL 图像中找到可能的外部边界框，函数会将图像转换为灰度图，并使用不同的阈值和方法来二值化图像，提取可能的边界框。
# 函数会尝试多种方法来提取边界框，例如自适应阈值、Otsu 阈值和 Canny 边缘检测，并将提取到的边界框进行过滤和比较，最终
# 返回一个最合适的边界框，确保能够找到图像中实际的图形区域，而不是被文本或其他元素干扰。
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
# 获取 PDF 页面中可能的边界框候选区域，函数会调用 get_page_struct_info() 来提取页面的结构信息，包括文本框、图像边界框和绘图对象边界框。
# 然后函数会将这些边界框从 PDF 坐标系转换为像素坐标系，并使用 mask_text_regions_on_pil() 来遮盖文本区域，最后调用 find_outer_bbox() 
# 来找到可能的图形区域边界框，确保能够获取到 PDF 页面中实际的图形区域边界框候选，以供后续使用，例如在裁切图像时参考这些边界框来确定裁切范围。
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
# 将边界框扩展一定的像素值，并确保扩展后的边界框仍然在图像范围内，函数会根据指定的扩展像素值来调整边界框的坐标，
# 并使用 clamp_bbox() 来确保扩展后的边界框不会超出图像的宽度和高度，确保在裁切图像时能够正确地扩展边界框，同时
# 避免超出图像范围导致错误。
def _expand_bbox_clamped(bbox, exp_px, img_w, img_h):
    x0, y0, x1, y1 = bbox
    return clamp_bbox(x0 - exp_px, y0 - exp_px, x1 + exp_px, y1 + exp_px, img_w, img_h)
# 根据参考边界框和 PDF 页面中的边界框候选区域来确定最终的裁切边界框，并将 PDF 页面渲染为 PIL 图像，最后裁切图
# 像并保存为 PNG 格式的字节内容返回，函数会根据参考边界框和 PDF 页面中的边界框候选区域来计算一个联合的边界框，
# 并在此基础上进行一定的扩展，确保裁切后的图像能够包含所有的图形内容，同时避免过多的空白区域。最后将裁切后的图
# 像保存为 PNG 格式的字节内容返回，供后续使用，例如嵌入到 PDF 中。
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
# 在给定宽度 w0 和高度 h0 的情况下，判断是否需要旋转图块以适应页面的宽度限制。函数会比较原始宽高和旋转后的宽高，
# 并根据是否满足页面宽度限制以及高度的比较来选择最佳的排布方式，确保在混拼布局时能够更紧凑地利用页面空间。
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
# 每行选一个 base_h行基准高度),同一类型可在该行做列×堆叠布局,通过行评分选择更紧凑的排法,高度跨过464mm时自动插入10mm缝,
# 最后再做一次hole fill,用剩余小图补空洞
def pack_mix_by_height_rule(mix_types, page_max_h=PAGE_H_MAX):
    X_START = float(PAGE_MARGIN_L)
    X_END = float(PAGE_W - PAGE_MARGIN_R)
    X_SPLIT = float(SPLIT_X)
    X_GAP_START = float(SPLIT_X + SPLIT_GAP_W)
    gap = float(CELL_GAP_MM)

    types_list = []
    for t in mix_types:
        tt = dict(t)
        tt["rem"] = int(tt.get("rem", tt.get("N", 0)) or 0)
        types_list.append(tt)

    for t in types_list:
        if int(t["rem"]) > 0 and float(t["W"]) > TOTAL_USABLE_W + 1e-9:
            print("⚠️ MIX_SKIP_TOO_WIDE_FOR_PAGE:", t.get("tid"), "W=%.2f > %.2f" % (float(t["W"]), TOTAL_USABLE_W))
            t["rem"] = 0

# 根据 mix_types 中的每个类型的高度来确定行基准高度，并在每行中尝试放置尽可能多的图块，直到达到页面的最大高度限制。
# 函数会根据行基准高度来选择适合的图块进行排布,并在每行结束后更新剩余的图块数量，最后返回一个包含每行布局信息的列表，
# 供后续使用，例如在生成 PDF 时参考这些布局信息来放置图块。
    def _build_one_row(work_list, base_h, y):
        x = X_START
        in_left = True
        row_blocks = []
        row_added = []

        while True:
            region_end = X_SPLIT if in_left else X_END
            rem_w = float(region_end) - float(x)

            if in_left and rem_w <= 1e-6:
                in_left = False
                x = X_GAP_START
                continue
            if (not in_left) and rem_w <= 1e-6:
                break

            cand = None
            cand_rank = None
            for t in work_list:
                if int(t["rem"]) <= 0:
                    continue
                if float(t["H"]) > base_h + 1e-9:
                    continue
                if float(t["W"]) > rem_w + 1e-9:
                    continue

                Hi = float(t["H"])
                Wi = float(t["W"])

                stack_cap = int(math.floor((base_h + gap) / (Hi + gap))) if Hi > 0 else 1
                if stack_cap <= 0:
                    stack_cap = 1

                max_cols = int(math.floor((rem_w + gap) / (Wi + gap))) if Wi > 0 else 0
                if max_cols <= 0:
                    continue

                need_cols = int(math.ceil(float(t["rem"]) / float(stack_cap)))
                cols = min(max_cols, need_cols)
                if cols <= 0:
                    cols = 1

                block_w = float(cols) * Wi + float(max(0, cols - 1)) * gap
                leftover = rem_w - block_w

                rank = (-Hi, float(leftover), -Wi)
                if cand is None or rank < cand_rank:
                    cand = t
                    cand_rank = rank

            if cand is None:
                if in_left:
                    in_left = False
                    x = X_GAP_START
                    continue
                break

            Hi = float(cand["H"])
            Wi = float(cand["W"])
            stack_cap = int(math.floor((base_h + gap) / (Hi + gap))) if Hi > 0 else 1
            if stack_cap <= 0:
                stack_cap = 1
            max_cols = int(math.floor((rem_w + gap) / (Wi + gap))) if Wi > 0 else 0
            if max_cols <= 0:
                if in_left:
                    in_left = False
                    x = X_GAP_START
                    continue
                break

            need_cols = int(math.ceil(float(cand["rem"]) / float(stack_cap)))
            cols = min(max_cols, need_cols)
            if cols <= 0:
                cols = 1

            x0 = float(x)
            block_w = float(cols) * Wi + float(max(0, cols - 1)) * gap
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
                        "rot90": bool(cand.get("rot90", False)),
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

            if in_left and x >= X_SPLIT - 1e-9:
                in_left = False
                x = X_GAP_START

            if (not in_left) and x >= X_END - 1e-6:
                break

        if not row_added:
            return [], [], 0.0, 0.0

        row_max_end = 0.0
        for p in row_added:
            row_max_end = max(row_max_end, float(p["y"]) + float(p["h"]))
        row_used_h = max(float(QR_BAND), row_max_end - float(y))
        return row_blocks, row_added, float(row_used_h), float(row_max_end)

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
        y = 0.0

        while True:
            types_list = [t for t in types_list if int(t["rem"]) > 0]
            if not types_list:
                break

            hs = sorted({float(t["H"]) for t in types_list}, reverse=True)
            cand_hs = hs[:6] if len(hs) > 6 else hs

            best = None
            best_score = None

            for bh in cand_hs:
                sim_list = [dict(t) for t in types_list]
                row_blocks, row_added, row_used_h, _row_max_end = _build_one_row(sim_list, bh, y)
                if not row_added:
                    continue

                span = False
                span_x1 = None
                left_max = X_START
                right_max = X_GAP_START
                for b in row_blocks:
                    x0 = float(b["x0"])
                    x1 = float(b["x1"])
                    if x0 < X_SPLIT - 1e-6 and x1 > X_GAP_START + 1e-6:
                        span = True
                        span_x1 = x1
                    if x0 < X_SPLIT + 1e-6:
                        left_max = max(left_max, x1)
                    if x0 >= X_GAP_START - 1e-6:
                        right_max = max(right_max, x1)

                if span:
                    leftover = max(0.0, X_END - float(span_x1 or left_max))
                else:
                    leftover_left = max(0.0, X_SPLIT - left_max)
                    leftover_right = max(0.0, X_END - right_max)
                    leftover = leftover_left + leftover_right

                used_area = 0.0
                for p in row_added:
                    used_area += float(p["w"]) * float(p["h"])
                denom = max(1.0, float(row_used_h) * float(TOTAL_USABLE_W))
                density = used_area / denom

                score = (-density, float(leftover), float(row_used_h))

                if best is None or score < best_score:
                    best = (bh, row_blocks, row_added, row_used_h, _row_max_end)
                    best_score = score

            if best is None:
                break

            bh, _rb, _ra, _ruh, _rme = best

            row_blocks2, row_added2, row_used_h2, row_max_end2 = _build_one_row(types_list, bh, y)
            if not row_added2:
                types_list.sort(key=lambda a: (-float(a.get("W", 0.0)), -float(a.get("H", 0.0))))
                types_list[0]["rem"] = 0
                continue

            page["blocks"].extend(row_blocks2)
            page["placements"].extend(row_added2)

            y2 = float(y) + float(row_used_h2)

            boundary = 464.0
            while boundary <= float(y) + 1e-9:
                boundary += 464.0
            if float(y) < boundary and y2 > boundary + 1e-9:
                y2 += 10.0

            y = y2
            page["used_h"] = max(page["used_h"], float(row_max_end2), float(y))

            if y >= float(page_max_h) - 1e-6:
                break

        if not page["placements"]:
            break

        _mix_hole_fill(page, types_list, page_max_h, gap=float(CELL_GAP_MM))
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
                cand = (fy, fx, fx, fy, fw, fh)
                if best is None or cand < best:
                    best = cand
        if best is None:
            return None
        return {"x": best[2], "y": best[3]}

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
    cand_types.sort(key=lambda t: -(float(t["W"]) * float(t["H"])))

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
            opts = [(w0, h0, bool(t.get("rot90", False)))]
            if abs(w0 - h0) > 1e-6:
                opts.append((h0, w0, (not bool(t.get("rot90", False)))))

            def _ok(w):
                return w <= LEFT_BLOCK_W + 1e-9

            opts.sort(key=lambda x: (0 if _ok(x[0]) else 1, x[0]))

            best_pick = None
            for (wi, hi, rot90) in opts:
                rw = int(round(wi))
                rh = int(round(hi))
                pos = bp.find_bottom_left(rw + int(round(gap)), rh + int(round(gap)))
                if pos is None:
                    continue
                best_pick = (pos, wi, hi, rot90)
                break

            if best_pick is None:
                continue

            pos, wi, hi, rot90 = best_pick
            bp.cut_out(pos["x"], pos["y"], int(round(wi + gap)), int(round(hi + gap)))

            page["placements"].append({
                "type": t,
                "tid": t["tid"],
                "x": float(pos["x"]) + usable_x0,
                "y": float(pos["y"]) + usable_y0,
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
        for p in page["placements"]:
            max_end = max(max_end, float(p["y"]) + float(p["h"]))
        page["used_h"] = max(page.get("used_h", 0.0), max_end)

def _mix_draw_edge_ticks(page, items, page_h_mm):
    if not items:
        return

    eps = 1.0  # mm 容差

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

    # 左侧块：画左侧刀线
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
        img_bytes = t["img_cont"] if is_contour else t["img_body"]

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

    # 最后画刀线
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
            outer_w0 = float(A + 2.0 * OUTER_EXT)
            outer_h0 = float(B + 2.0 * OUTER_EXT)

            if int(N) >= MIX_THRESHOLD:
                skip_list.append((type_name_raw, "skip_N_ge_%d" % MIX_THRESHOLD))
                _log(log_cb, "⏭️ SKIP(N>=%d): %s" % (MIX_THRESHOLD, type_name_raw))
                continue

            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)
            qr_bytes = make_qr_png_bytes(qr_text)

            for pi in range(pair_count):
                page_body = 2 * pi
                page_cont = 2 * pi + 1

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

                Mi, Ni, rot90 = _mix_choose_orient(outer_w0, outer_h0)
                mix_types.append({
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

            doc1 = render_mix_page(pg, is_contour=False)
            doc2 = render_mix_page(pg, is_contour=True)

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