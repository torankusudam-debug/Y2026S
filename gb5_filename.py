# -*- coding: utf-8 -*-
import os
import re


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
            raise ValueError("无法从文件名解析 A×B 与 N（新规则）：%s" % base)

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