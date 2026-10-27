#!/usr/bin/env python3
"""DOCGEN pdf text: replace text inside an existing PDF without re-creating
it. Only the affected line/paragraph is rewritten; the edit is verified by
reading the page back (other lines and pages must be unchanged).

Usage:
    pdf_text.py <in.pdf> <out.pdf> "<old text>" "<new text>"
                [--page N] [--near "row key"] [--occurrence N] [--band header|footer]
                [--strategy auto|inplace|redraw]

Narrow ambiguous targets with --page / --near (e.g. the product name of a
table row) / --occurrence. Refusals (exit 1) are honest limits: the line mixes
formatting, the font family is not installed, or the text does not fit.
Changing color/bold/font of PDF text is not supported (edit the source file).
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
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--page", type=int)
    ap.add_argument("--near")
    ap.add_argument("--occurrence", type=int)
    ap.add_argument("--band", choices=["header", "footer", "any"])
    ap.add_argument("--strategy", choices=["auto", "inplace", "redraw"], default="auto")
    a = ap.parse_args()
    try:
        rep = F.replace_text(a.src, a.dst, a.old, a.new, page=a.page, occurrence=a.occurrence,
                             band=a.band, near=a.near, strategy=a.strategy)
    except DocgenError as e:
        print(f"[pdf_text] REFUSED  {e}")
        return 1
    print(f"[pdf_text] OK  {json.dumps(rep, ensure_ascii=False)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
