# -*- coding: utf-8 -*-
import os
from pathlib import Path

IN_DIR = Path(r"D:\test_data\iest")
RECURSIVE = True
OVERWRITE = False  # True: 目标.ai存在就覆盖；False: 跳过

def main():
    if not IN_DIR.exists():
        print("[FATAL] not found:", IN_DIR)
        return

    files = IN_DIR.rglob("*.pdf") if RECURSIVE else IN_DIR.glob("*.pdf")
    cnt_ok = cnt_skip = cnt_err = 0

    for pdf in files:
        try:
            ai = pdf.with_suffix(".ai")
            if ai.exists():
                if OVERWRITE:
                    ai.unlink()
                else:
                    print("[SKIP exists]", ai)
                    cnt_skip += 1
                    continue
            pdf.rename(ai)
            print("[OK]", pdf, "->", ai)
            cnt_ok += 1
        except Exception as e:
            print("[ERR]", pdf, e)
            cnt_err += 1

    # 也处理 .PDF
    files2 = IN_DIR.rglob("*.PDF") if RECURSIVE else IN_DIR.glob("*.PDF")
    for pdf in files2:
        try:
            ai = pdf.with_suffix(".ai")
            if ai.exists():
                if OVERWRITE:
                    ai.unlink()
                else:
                    print("[SKIP exists]", ai)
                    cnt_skip += 1
                    continue
            pdf.rename(ai)
            print("[OK]", pdf, "->", ai)
            cnt_ok += 1
        except Exception as e:
            print("[ERR]", pdf, e)
            cnt_err += 1

    print(f"[DONE] OK={cnt_ok} SKIP={cnt_skip} ERR={cnt_err}")

if __name__ == "__main__":
    main()