#!/usr/bin/env python3
"""DOCGEN legacy conversion: .doc/.xls/.ppt (and .rtf/.odt/.ods/.odp) ->
.docx/.xlsx/.pptx with LibreOffice, so the result can be edited with the
docgen helpers. The conversion itself may change layout details: always
tell the user the file was converted and deliver the new format.

Usage:
    convert_legacy.py <file> [--outdir /workspace/scratch]
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from docgen import DocgenError, office  # noqa: E402

TARGET = {".doc": "docx", ".rtf": "docx", ".odt": "docx", ".xls": "xlsx", ".ods": "xlsx",
          ".ppt": "pptx", ".odp": "pptx"}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--outdir", default="/workspace/scratch")
    a = ap.parse_args()
    src = Path(a.file)
    fmt = TARGET.get(src.suffix.lower())
    if not src.exists() or not fmt:
        print(f"usage error: pass an existing {', '.join(TARGET)} file", file=sys.stderr)
        return 2
    try:
        out = office.convert(src, fmt, a.outdir)
    except DocgenError as e:
        print(f"[convert_legacy] FAIL  {e}")
        return 1
    print(f"[convert_legacy] OK  {src.name} -> {out}")
    print("  note for the final answer: the file was converted to "
          f".{fmt}; layout details may differ from the original")
    return 0


if __name__ == "__main__":
    sys.exit(main())
