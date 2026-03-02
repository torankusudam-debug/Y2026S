# -*- coding: utf-8 -*-
"""
gb5_merge.py  —— 垂直连接合并工具

功能：
  将多个“单页临时PDF”按顺序垂直拼接到一张纸上。
  宽度统一为 MERGE_W（600mm），高度累加不超过 MERGE_H_MAX（1080mm）。
  超过时开始新的一张纸。每个输出PDF只有一页。

使用方式：
  pages_info = [
      {"doc": fitz_doc, "page_index": 0, "W_mm": 600, "H_mm": 350},
      {"doc": fitz_doc, "page_index": 0, "W_mm": 600, "H_mm": 400},
      ...
  ]
  merged_docs = merge_pages_vertical(pages_info, max_w_mm=600, max_h_mm=1080)
  # 返回 list of fitz.Document，每个只有一页
"""

import fitz
from gb5_utils import mm_to_pt

# 默认参数
MERGE_W = 600.0
MERGE_H_MAX = 1080.0


def merge_pages_vertical(pages_info, max_w_mm=MERGE_W, max_h_mm=MERGE_H_MAX):
    """
    将多个单页PDF垂直拼接。

    参数:
        pages_info: list of dict，每个dict包含:
            - "doc": fitz.Document (已打开的临时文档)
            - "page_index": int (页面索引，通常为0)
            - "W_mm": float (该页宽度mm)
            - "H_mm": float (该页高度mm)
        max_w_mm: 输出纸张宽度(mm)
        max_h_mm: 输出纸张最大高度(mm)

    返回:
        list of fitz.Document，每个文档只有一页
    """
    if not pages_info:
        return []

    # ---- 分组：按高度限制分组 ----
    groups = []        # 每组是一个 list of pages_info items
    cur_group = []
    cur_h = 0.0

    for pi in pages_info:
        h = pi["H_mm"]
        if cur_group and (cur_h + h > max_h_mm):
            # 当前组已满，开始新组
            groups.append(cur_group)
            cur_group = [pi]
            cur_h = h
        else:
            cur_group.append(pi)
            cur_h += h

    if cur_group:
        groups.append(cur_group)

    # ---- 每组合并为一个单页PDF ----
    result_docs = []
    for group in groups:
        total_h = sum(item["H_mm"] for item in group)
        out_w_pt = mm_to_pt(max_w_mm)
        out_h_pt = mm_to_pt(total_h)

        out_doc = fitz.open()
        out_page = out_doc.new_page(width=out_w_pt, height=out_h_pt)

        y_offset_pt = 0.0
        for item in group:
            src_doc = item["doc"]
            src_page_index = item["page_index"]
            src_page = src_doc.load_page(src_page_index)
            src_rect = src_page.rect  # 源页面的矩形

            w_mm = item["W_mm"]
            h_mm = item["H_mm"]
            w_pt = mm_to_pt(w_mm)
            h_pt = mm_to_pt(h_mm)

            # 目标矩形：水平居左（宽度可能一致），垂直按偏移
            dst_rect = fitz.Rect(0, y_offset_pt, w_pt, y_offset_pt + h_pt)

            # 将源页面作为 XObject 插入到目标页面
            out_page.show_pdf_page(dst_rect, src_doc, src_page_index)

            y_offset_pt += h_pt

        # 绘制整张纸的外框
        out_page.draw_rect(fitz.Rect(0, 0, out_w_pt, out_h_pt), color=(0, 0, 0), width=1.0)

        result_docs.append(out_doc)

    return result_docs
