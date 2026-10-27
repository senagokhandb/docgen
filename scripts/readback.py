#!/usr/bin/env python3
"""DOCGEN read-back: print the content of an office or PDF file.

Usage:
    readback.py <file> [--max-chars N] [--tables]

- .docx/.xlsx/.pptx (+ legacy .doc/.xls/.ppt): office_oxide -> markdown
- .pdf: pdfplumber per-page text (and table summaries with --tables)
"""
import argparse
import sys
from pathlib import Path

OFFICE_EXTS = {".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"}


def read_office(path: Path) -> str:
    from office_oxide import Document
    doc = Document.open(str(path))
    try:
        return doc.to_markdown()
    finally:
        doc.close()


def read_pdf(path: Path, with_tables: bool) -> str:
    import pdfplumber
    out = []
    with pdfplumber.open(str(path)) as pdf:
        out.append(f"(pdf pages: {len(pdf.pages)})")
        for i, page in enumerate(pdf.pages, 1):
            out.append(f"\n--- page {i} ---")
            text = page.extract_text() or "(no text layer)"
            out.append(text)
            if with_tables:
                tables = page.extract_tables()
                for j, t in enumerate(tables, 1):
                    out.append(f"[table {j}] {len(t)} rows x {len(t[0]) if t else 0} cols")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("file")
    ap.add_argument("--max-chars", type=int, default=20000,
                    help="truncate output (default 20000)")
    ap.add_argument("--tables", action="store_true",
                    help="include table summaries for PDFs")
    args = ap.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"file not found: {path}", file=sys.stderr)
        return 2

    try:
        if path.suffix.lower() == ".pdf":
            content = read_pdf(path, args.tables)
        elif path.suffix.lower() in OFFICE_EXTS:
            content = read_office(path)
        else:
            print(f"unsupported extension: {path.suffix}", file=sys.stderr)
            return 2
    except Exception as e:
        print(f"read-back failed: {e}", file=sys.stderr)
        return 1

    if len(content) > args.max_chars:
        content = content[: args.max_chars] + f"\n... [truncated at {args.max_chars} chars]"
    print(content)
    return 0


if __name__ == "__main__":
    sys.exit(main())
