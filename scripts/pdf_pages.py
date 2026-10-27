#!/usr/bin/env python3
"""DOCGEN pdf pages: page-level edits that copy pages unchanged (pdfium).

Usage:
    pdf_pages.py <in.pdf> <out.pdf> delete 4            # or "2,5-6"
    pdf_pages.py <in.pdf> <out.pdf> keep 1-3
    pdf_pages.py <in.pdf> <out.pdf> move 5 --before 4   # original numbering; --before N+1 = to the end
    pdf_pages.py <in.pdf> <out.pdf> insert other.pdf [--at 3] [--pages 1-2]   # default: append
    pdf_pages.py <in.pdf> <out.pdf> rotate 2 --degrees 90
Add --renumber to rewrite "n / N" page numbers in the header/footer band.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from docgen import DocgenError  # noqa: E402
from docgen import pdf as F  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("op", choices=["delete", "keep", "move", "insert", "rotate"])
    ap.add_argument("arg", help="pages (delete/keep/move/rotate) or the PDF to insert")
    ap.add_argument("--before", type=int)
    ap.add_argument("--at", type=int)
    ap.add_argument("--pages")
    ap.add_argument("--degrees", type=int, default=90)
    ap.add_argument("--renumber", action="store_true")
    a = ap.parse_args()
    try:
        if a.op == "delete":
            info = {"pages": F.delete_pages(a.src, a.dst, a.arg)}
        elif a.op == "keep":
            info = {"pages": F.keep_pages(a.src, a.dst, a.arg)}
        elif a.op == "move":
            if a.before is None:
                print("usage error: move needs --before", file=sys.stderr)
                return 2
            info = {"order": F.move_pages(a.src, a.dst, a.arg, a.before)}
        elif a.op == "insert":
            info = {"pages": F.insert_pdf(a.src, a.dst, a.arg, at=a.at, pages=a.pages)}
        else:
            F.rotate_pages(a.src, a.dst, a.arg, a.degrees)
            info = {"rotated": a.arg}
        if a.renumber:
            info["renumber"] = F.renumber(a.dst, a.dst)
    except DocgenError as e:
        print(f"[pdf_pages] FAIL  {e}")
        return 1
    print(f"[pdf_pages] OK  {a.op} -> {a.dst}  {json.dumps(info, ensure_ascii=False)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
