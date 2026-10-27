#!/usr/bin/env python3
"""DOCGEN find target: list every place a text occurs, with enough context
to pick the right one, BEFORE editing.

Usage:
    find_target.py <file> "<text>" [--section S] [--slide N] [--page N] [--sheet S]
                   [--whole-word] [--ignore-case] [--headings] [--band header|footer] [--json]

Exit codes: 0 = exactly one match, 1 = no match, 3 = several matches
(ambiguous: narrow it down or report the candidates to the user), 2 = usage.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from docgen.compare import kind_of  # noqa: E402
from docgen.package import Package  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("text")
    ap.add_argument("--section")
    ap.add_argument("--slide", type=int)
    ap.add_argument("--page", type=int)
    ap.add_argument("--sheet")
    ap.add_argument("--band", choices=["header", "footer", "any"])
    ap.add_argument("--whole-word", action="store_true")
    ap.add_argument("--ignore-case", action="store_true")
    ap.add_argument("--headings", action="store_true", help="DOCX: match heading paragraphs only")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    kind = kind_of(a.file)
    if kind is None or not Path(a.file).exists():
        print("usage error: pass an existing .docx/.pptx/.xlsx/.pdf", file=sys.stderr)
        return 2
    if kind == "docx":
        from docgen import docx as D
        pkg = Package(a.file)
        hits = [repr(m) for m in D.find(pkg, a.text, section=a.section, whole_word=a.whole_word,
                                        ignore_case=a.ignore_case, headings_only=a.headings)]
    elif kind == "pptx":
        from docgen import pptx as P
        pkg = Package(a.file)
        hits = [repr(m) for m in P.find(pkg, a.text, slide=a.slide, whole_word=a.whole_word, ignore_case=a.ignore_case)]
        hits += [f"slide {s.index} notes" for s in P.slides(pkg)
                 if (a.slide in (None, s.index)) and a.text in P.get_notes(pkg, s.index)]
    elif kind == "xlsx":
        from docgen import xlsx as X
        pkg = Package(a.file)
        hits = [f"{ref} = {X.get(pkg, ref)[0]!r}" for ref in X.find(pkg, a.text, sheet=a.sheet, ignore_case=a.ignore_case)]
    else:
        from docgen import pdf as F
        hits = [repr(m) for m in F.find(a.file, a.text, page=a.page, band=a.band)]

    if a.json:
        print(json.dumps({"file": a.file, "text": a.text, "count": len(hits), "matches": hits}, ensure_ascii=False, indent=2))
    else:
        print(f"[find_target] {len(hits)} match(es) for {a.text!r} in {a.file}")
        for i, h in enumerate(hits, 1):
            print(f"  {i}. {h}")
    return 0 if len(hits) == 1 else (1 if not hits else 3)


if __name__ == "__main__":
    sys.exit(main())
