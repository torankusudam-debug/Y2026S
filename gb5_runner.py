# -*- coding: utf-8 -*-
import os
import fitz

import gb5_config as C
from gb5_utils import ensure_dir, _log, archive_input_pdf_to_dir, _sanitize_filename_for_windows, safe_save, union_bbox
from gb5_filename import parse_A_B_N_from_filename, extract_label_text, extract_qr_text_from_filename
from gb5_qr import make_qr_png_bytes
from gb5_bbox import get_page_bbox_candidates_px, make_part_png_bytes_using_ref_bbox
from gb5_pdf_struct import render_page_to_pil
from gb5_single import solve_single_type_no_waste, build_single_placements, append_pages_single
from gb5_full import pack_full_sequential_sheets, append_one_full_sheet


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

    N_total = len(archived_paths)

    for idx, path in enumerate(archived_paths, start=1):
        type_name_raw = os.path.splitext(os.path.basename(path))[0]
        type_name = _sanitize_filename_for_windows(type_name_raw)

        if progress_cb:
            progress_cb(idx - 1, N_total, "处理中 %d/%d  %s" % (idx - 1, N_total, type_name_raw))

        d = None
        out1_single = fitz.open()
        out2_single = fitz.open()
        has_single_pages = False

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

            for pi in range(pair_count):
                page_body = 2 * pi
                page_cont = 2 * pi + 1

                best = solve_single_type_no_waste(outer_w, outer_h, N)

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

                if best is None:
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

                append_pages_single(out1_single, best, placements, best["pages"], qr_bytes, label_text, img_body)
                append_pages_single(out2_single, best, placements, best["pages"], qr_bytes, label_text, img_cont)

                has_single_pages = True
                ok_list.append((tid, best["pages"]))
                _log(log_cb, "✅ 整拼OK: %s pages=%d W=%.1f H=%.1f" % (tid, best["pages"], best["W"], best["H"]))

        except Exception as e:
            skip_list.append((type_name_raw, repr(e)))
            _log(log_cb, "⚠️ SKIP: %s reason=%s" % (type_name_raw, repr(e)))
        finally:
            try:
                if d is not None:
                    d.close()
            except Exception:
                pass

            if has_single_pages:
                out_path1 = os.path.join(C.DEST_DIR1, type_name + ".pdf")
                out_path2 = os.path.join(C.DEST_DIR2, type_name + ".pdf")
                p1 = safe_save(out1_single, out_path1)
                p2 = safe_save(out2_single, out_path2)
                single_outputs_p1.append(p1)
                single_outputs_p2.append(p2)
                _log(log_cb, "📄 整拼保存：%s | %s" % (p1, p2))
            else:
                try:
                    out1_single.close()
                except Exception:
                    pass
                try:
                    out2_single.close()
                except Exception:
                    pass

    full_p1 = None
    full_p2 = None
    if fullmix_pool:
        _log(log_cb, "🧩 开始顺序全拼（汇总到同一个PDF）：pool_types=%d" % len(fullmix_pool))

        sheets, by_tid = pack_full_sequential_sheets(fullmix_pool, log_cb=log_cb)
        out1_full = fitz.open()
        out2_full = fitz.open()

        for si, sh in enumerate(sheets, start=1):
            append_one_full_sheet(out1_full, sh, by_tid, is_contour=False)
            append_one_full_sheet(out2_full, sh, by_tid, is_contour=True)
            ok_list.append(("FULLSEQ#%d(owner=%s)" % (si, sh["owner_tid"]), 1))

        full_path1 = os.path.join(C.DEST_DIR1, u"全拼.pdf")
        full_path2 = os.path.join(C.DEST_DIR2, u"全拼.pdf")
        full_p1 = safe_save(out1_full, full_path1)
        full_p2 = safe_save(out2_full, full_path2)
        _log(log_cb, "📌 全拼保存：%s | %s" % (full_p1, full_p2))

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