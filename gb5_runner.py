# -*- coding: utf-8 -*-
"""
gb5_runner.py —— 拼版主流程

核心逻辑：
  1. 遍历所有输入PDF，解析文件名获取尺寸、数量、版号
  2. N < FULL_THRESHOLD (10) → 直接进入全拼
  3. N >= 10 → 尝试整拼，整拼无解则进入全拼
  4. 整拼结果按版号分组，相同版号的垂直拼接到同一个PDF
  5. 全拼结果垂直拼接，命名为 全拼_i
  6. 每个PDF只有一页，高度不超过1500mm
"""
import os
from collections import OrderedDict

import fitz

import gb5_config as C
from gb5_utils import ensure_dir, _log, archive_input_pdf_to_dir, _sanitize_filename_for_windows, safe_save, union_bbox
from gb5_filename import parse_A_B_N_from_filename, extract_label_text, extract_qr_text_from_filename, extract_ban_hao
from gb5_qr import make_qr_png_bytes
from gb5_bbox import get_page_bbox_candidates_px, make_part_png_bytes_using_ref_bbox
from gb5_pdf_struct import render_page_to_pil
from gb5_single import solve_single_type_no_waste, build_single_placements, append_pages_single
from gb5_full import pack_full_grid_sheets, append_one_full_sheet
from gb5_merge import merge_pages_vertical


def _build_single_temp_docs(best, placements, pages, qr_bytes, label_text, img_bytes):
    """
    为整拼生成临时的单页fitz文档列表（每页一个doc），用于后续垂直合并。
    返回 list of dict: [{"doc": fitz.Document, "page_index": 0, "W_mm": W, "H_mm": H}, ...]
    """
    W = best["W"]
    H = best["H"]
    result = []
    for _ in range(pages):
        tmp_doc = fitz.open()
        append_pages_single(tmp_doc, best, placements, 1, qr_bytes, label_text, img_bytes)
        result.append({
            "doc": tmp_doc,
            "page_index": 0,
            "W_mm": W,
            "H_mm": H,
        })
    return result


def _build_full_temp_doc(sheet, by_tid, is_contour):
    """
    为全拼的一个sheet生成临时的单页fitz文档，用于后续垂直合并。
    返回 dict: {"doc": fitz.Document, "page_index": 0, "W_mm": W, "H_mm": H}
    """
    tmp_doc = fitz.open()
    append_one_full_sheet(tmp_doc, sheet, by_tid, is_contour=is_contour)
    return {
        "doc": tmp_doc,
        "page_index": 0,
        "W_mm": int(sheet["W"]),
        "H_mm": int(sheet["H"]),
    }


def run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None):
    if cfg:
        C.DEST_DIR = cfg.get("DEST_DIR", C.DEST_DIR)
        C.IN_PDF_ARCHIVE_DIR = cfg.get("TEST_DIR", C.IN_PDF_ARCHIVE_DIR)
        C.DEST_DIR1 = cfg.get("DEST_DIR1", C.DEST_DIR1)
        C.DEST_DIR2 = cfg.get("DEST_DIR2", C.DEST_DIR2)
        C.OUT_PDF_P1 = os.path.join(C.DEST_DIR1, "over_test_p1.pdf")
        C.OUT_PDF_P2 = os.path.join(C.DEST_DIR2, "over_test_p2.pdf")

    ensure_dir(C.IN_PDF_ARCHIVE_DIR)
    ensure_dir(C.DEST_DIR1)
    ensure_dir(C.DEST_DIR2)

    if input_pdfs is not None:
        archived_paths = list(input_pdfs)
    else:
        pdfs = []
        for fn in os.listdir(C.DEST_DIR):
            lfn = fn.lower()
            if (not lfn.endswith(".pdf")) or lfn.startswith("over_test"):
                continue
            pdfs.append(os.path.join(C.DEST_DIR, fn))
        pdfs.sort()

        archived_paths = []
        for src in pdfs:
            try:
                dst = archive_input_pdf_to_dir(src, C.IN_PDF_ARCHIVE_DIR)
                archived_paths.append(dst)
            except Exception as e:
                _log(log_cb, "⚠️ 复制到test失败：%s err=%s" % (os.path.basename(src), repr(e)))

    if not archived_paths:
        raise RuntimeError("没有可处理PDF（archived_paths为空）")

    single_outputs_p1 = []
    single_outputs_p2 = []
    fullmix_pool = []
    ok_list = []
    skip_list = []

    # ---- 整拼临时页面按版号分组 ----
    # 结构: ban_hao_groups = OrderedDict {
    #   ban_hao: {
    #       "type_name": str (第一个源文件名，用于单文件命名),
    #       "pages_p1": [page_info, ...],
    #       "pages_p2": [page_info, ...],
    #       "file_count": int (该版号下有多少个源文件),
    #   }
    # }
    ban_hao_groups = OrderedDict()

    N_total = len(archived_paths)

    for idx, path in enumerate(archived_paths, start=1):
        type_name_raw = os.path.splitext(os.path.basename(path))[0]
        type_name = _sanitize_filename_for_windows(type_name_raw)

        if progress_cb:
            progress_cb(idx - 1, N_total, "处理中 %d/%d  %s" % (idx - 1, N_total, type_name_raw))

        d = None
        temp_pages_p1 = []
        temp_pages_p2 = []

        try:
            d = fitz.open(path)
            pc = d.page_count
            if pc < 2:
                raise RuntimeError("PDF页数不足2页")

            pair_count = pc // 2
            if pair_count <= 0:
                raise RuntimeError("PDF无有效页对(页数=%d)" % pc)

            if (pc % 2) == 1:
                _log(log_cb, "⚠️ %s 页数为奇数(%d)，最后一页将忽略" % (type_name_raw, pc))

            A, B, N = parse_A_B_N_from_filename(path)
            outer_w = int(round(A + 2.0 * C.OUTER_EXT))
            outer_h = int(round(B + 2.0 * C.OUTER_EXT))

            label_text = extract_label_text(path)
            qr_text = extract_qr_text_from_filename(path)
            ban_hao = extract_ban_hao(path)

            # ---- 判断是否直接全拼 ----
            force_full = (N < C.FULL_THRESHOLD)
            if force_full:
                _log(log_cb, "📋 N=%d < %d，直接进入全拼：%s" % (N, C.FULL_THRESHOLD, type_name_raw))

            for pi in range(pair_count):
                page_body = 2 * pi
                page_cont = 2 * pi + 1

                ref_img = render_page_to_pil(path, page_index=page_cont, dpi=C.RENDER_DPI, doc=d)
                draw_px, img_px, cv_px = get_page_bbox_candidates_px(d, page_cont, ref_img)
                if draw_px is not None:
                    ref_bbox = union_bbox(draw_px, cv_px)
                elif img_px is not None:
                    ref_bbox = union_bbox(img_px, cv_px)
                else:
                    ref_bbox = cv_px
                ref_size = ref_img.size

                img_body = make_part_png_bytes_using_ref_bbox(path, page_body, ref_bbox, ref_size, dpi=C.RENDER_DPI, doc=d)
                img_cont = make_part_png_bytes_using_ref_bbox(path, page_cont, ref_bbox, ref_size, dpi=C.RENDER_DPI, doc=d)

                tid = "%s@P%d" % (type_name_raw, pi + 1)

                if force_full:
                    # N < 10，直接全拼
                    fullmix_pool.append({
                        "tid": tid, "pdf": path, "pair": pi,
                        "need": int(N), "ow": int(outer_w), "oh": int(outer_h),
                        "img_body": img_body, "img_cont": img_cont,
                        "label": label_text, "qr_text": qr_text,
                    })
                    continue

                # N >= 10，尝试整拼
                best = solve_single_type_no_waste(outer_w, outer_h, N)

                if best is None:
                    # 整拼无解，进入全拼
                    fullmix_pool.append({
                        "tid": tid, "pdf": path, "pair": pi,
                        "need": int(N), "ow": int(outer_w), "oh": int(outer_h),
                        "img_body": img_body, "img_cont": img_cont,
                        "label": label_text, "qr_text": qr_text,
                    })
                    _log(log_cb, "⚠️ 整拼无解 -> 进入全拼池: %s" % tid)
                    continue

                placements = build_single_placements(tid, best)
                qr_bytes = make_qr_png_bytes(qr_text)

                # 生成临时单页文档
                tmp_p1 = _build_single_temp_docs(best, placements, best["pages"], qr_bytes, label_text, img_body)
                tmp_p2 = _build_single_temp_docs(best, placements, best["pages"], qr_bytes, label_text, img_cont)

                temp_pages_p1.extend(tmp_p1)
                temp_pages_p2.extend(tmp_p2)

                ok_list.append((tid, best["pages"]))
                _log(log_cb, "✅ 整拼OK: %s pages=%d W=%.1f H=%.1f ban_hao=%s" %
                     (tid, best["pages"], best["W"], best["H"], ban_hao))

        except Exception as e:
            skip_list.append((type_name_raw, repr(e)))
            _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, repr(e)))
        finally:
            try:
                if d is not None:
                    d.close()
            except Exception:
                pass

            # 按版号分组收集整拼临时页面
            if temp_pages_p1:
                if ban_hao not in ban_hao_groups:
                    ban_hao_groups[ban_hao] = {
                        "type_name": type_name,
                        "pages_p1": [],
                        "pages_p2": [],
                        "file_count": 0,
                    }
                ban_hao_groups[ban_hao]["pages_p1"].extend(temp_pages_p1)
                ban_hao_groups[ban_hao]["pages_p2"].extend(temp_pages_p2)
                ban_hao_groups[ban_hao]["file_count"] += 1

    # ========================================================
    # 整拼垂直合并：按版号分组，相同版号拼接到同一个PDF
    # ========================================================
    zhengpin_counter = 0

    for ban_hao, grp in ban_hao_groups.items():
        type_name = grp["type_name"]
        pages_p1 = grp["pages_p1"]
        pages_p2 = grp["pages_p2"]
        file_count = grp["file_count"]

        # 垂直合并
        merged_p1 = merge_pages_vertical(pages_p1, max_w_mm=C.SINGLE_W_MAX, max_h_mm=C.SINGLE_H_MAX)
        merged_p2 = merge_pages_vertical(pages_p2, max_w_mm=C.SINGLE_W_MAX, max_h_mm=C.SINGLE_H_MAX)

        num_merged = len(merged_p1)

        if file_count == 1 and num_merged == 1:
            # 单个源文件且只产生一个PDF → 用源文件名
            out_path1 = os.path.join(C.DEST_DIR1, type_name + ".pdf")
            out_path2 = os.path.join(C.DEST_DIR2, type_name + ".pdf")
            p1 = safe_save(merged_p1[0], out_path1)
            p2 = safe_save(merged_p2[0], out_path2)
            single_outputs_p1.append(p1)
            single_outputs_p2.append(p2)
            _log(log_cb, "📄 整拼保存（源文件名）：%s" % p1)
        else:
            # 多个源文件或多个PDF → 用"整拼n"命名
            for mi in range(num_merged):
                zhengpin_counter += 1
                name = u"整拼%d" % zhengpin_counter
                out_path1 = os.path.join(C.DEST_DIR1, name + ".pdf")
                out_path2 = os.path.join(C.DEST_DIR2, name + ".pdf")
                p1 = safe_save(merged_p1[mi], out_path1)
                p2 = safe_save(merged_p2[mi], out_path2)
                single_outputs_p1.append(p1)
                single_outputs_p2.append(p2)
                _log(log_cb, "📄 整拼保存（整拼%d，版号=%s）：%s" % (zhengpin_counter, ban_hao, p1))

        # 关闭临时文档
        for item in pages_p1:
            try:
                item["doc"].close()
            except Exception:
                pass
        for item in pages_p2:
            try:
                item["doc"].close()
            except Exception:
                pass

    # ========================================================
    # 全拼部分：网格排列 → 生成临时单页 → 垂直合并 → 命名全拼_i
    # ========================================================
    full_p1 = None
    full_p2 = None
    full_outputs_p1 = []
    full_outputs_p2 = []

    if fullmix_pool:
        _log(log_cb, "🧩 开始网格全拼：pool_types=%d" % len(fullmix_pool))

        sheets, by_tid = pack_full_grid_sheets(fullmix_pool, log_cb=log_cb)

        # 为每个sheet生成临时单页文档
        temp_full_p1 = []
        temp_full_p2 = []
        for si, sh in enumerate(sheets, start=1):
            tmp1 = _build_full_temp_doc(sh, by_tid, is_contour=False)
            tmp2 = _build_full_temp_doc(sh, by_tid, is_contour=True)
            temp_full_p1.append(tmp1)
            temp_full_p2.append(tmp2)
            ok_list.append(("FULLGRID#%d(owner=%s)" % (si, sh["owner_tid"]), 1))

        # 垂直合并全拼
        merged_full_p1 = merge_pages_vertical(temp_full_p1, max_w_mm=C.FULL_W_MAX, max_h_mm=C.FULL_H_MAX)
        merged_full_p2 = merge_pages_vertical(temp_full_p2, max_w_mm=C.FULL_W_MAX, max_h_mm=C.FULL_H_MAX)

        num_full = len(merged_full_p1)
        for fi in range(num_full):
            name = u"全拼_%d" % (fi + 1)
            full_path1 = os.path.join(C.DEST_DIR1, name + ".pdf")
            full_path2 = os.path.join(C.DEST_DIR2, name + ".pdf")
            fp1 = safe_save(merged_full_p1[fi], full_path1)
            fp2 = safe_save(merged_full_p2[fi], full_path2)
            full_outputs_p1.append(fp1)
            full_outputs_p2.append(fp2)
            _log(log_cb, "📌 全拼保存（全拼_%d）：%s" % (fi + 1, fp1))

        full_p1 = full_outputs_p1[0] if full_outputs_p1 else None
        full_p2 = full_outputs_p2[0] if full_outputs_p2 else None

        # 关闭临时文档
        for item in temp_full_p1:
            try:
                item["doc"].close()
            except Exception:
                pass
        for item in temp_full_p2:
            try:
                item["doc"].close()
            except Exception:
                pass

    if progress_cb:
        progress_cb(N_total, N_total, "完成 %d/%d" % (N_total, N_total))

    return {
        "single_p1_files": single_outputs_p1,
        "single_p2_files": single_outputs_p2,
        "full_p1": full_p1,
        "full_p2": full_p2,
        "ok": ok_list,
        "skip": skip_list,
    }


def main():
    res = run(cfg=None, input_pdfs=None, progress_cb=None, log_cb=None)
    print("DONE:")
    print("  single_p1_files:", len(res.get("single_p1_files") or []))
    print("  single_p2_files:", len(res.get("single_p2_files") or []))
    print("  full_p1:", res.get("full_p1"))
    print("  full_p2:", res.get("full_p2"))


if __name__ == "__main__":
    main()
