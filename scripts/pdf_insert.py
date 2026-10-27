#!/usr/bin/env python3
"""DOCGEN pdf insert: add new paragraphs to an existing PDF page, right
after a given line, in the page's body font/size/color and text column.
Nothing else on the page or in the document moves. If the text does not fit
before the next block or the footer, the edit is refused (exit 1): ask the
user where it should go — e.g. render new pages with HTML and add them with
pdf_pages.py insert.

Usage:
    pdf_insert.py <in.pdf> <out.pdf> --after "<text of the line>" --text "<paragraph>" [--text ...]
                  [--page N] [--near "context"]
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
    ap.add_argument("--after", required=True)
    ap.add_argument("--text", action="append", required=True, help="one paragraph (repeatable)")
    ap.add_argument("--page", type=int)
    ap.add_argument("--near")
    a = ap.parse_args()
    try:
        rep = F.insert_paragraphs(a.src, a.dst, a.after, a.text, page=a.page, near=a.near)
    except DocgenError as e:
        print(f"[pdf_insert] REFUSED  {e}")
        return 1
    print(f"[pdf_insert] OK  {json.dumps(rep, ensure_ascii=False)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
