#!/usr/bin/env python3
"""DOCGEN diff check: compare an edited file with its original and FAIL on
any change outside the declared scope (mandatory after editing an uploaded
file). Works for .docx .pptx .xlsx .pdf (and .dotx/.potx/.xltx).

Usage:
    diff_check.py <original> <result> [scope...] [--expect-text T] [--json] [--record manifest.json]

Scope (repeatable; declare exactly what the user asked to change):
    --allow-text T        paragraphs/shapes/lines/cells whose old or new text contains T
    --allow-section S     DOCX: everything under heading S ("Bulgular > Sonuç")
    --allow-style ID      DOCX: changes to style ID in styles.xml
    --allow-slide N       PPTX: anything on slide N (original numbering)
    --allow-notes N       PPTX: speaker notes of slide N
    --allow-page N        PDF: page N (original numbering)
    --allow-band          PDF: header/footer lines on any page
    --allow-structure     slides/pages added, removed, moved; sheets added/removed
    --allow-cell R        XLSX: 'Sheet!A1' or 'Sheet!A1:C9'
    --allow-sheet S       XLSX: anything on sheet S
    --allow-recalc        XLSX: cached formula results may change
    --allow-feature F     XLSX: cf|dv|merged|cols|tables|state|names|charts|drawings|ext[:Sheet]
    --allow-theme         theme parts (theme colors/fonts)
    --allow-part GLOB     any other package part, e.g. 'word/numbering.xml'
Kind of change (use one when the request is purely of that kind):
    --text-only           formatting must not change anywhere (text edits)
    --format-only         text must not change anywhere (color/font/size/... edits)
Expectations:
    --expect-text T       T must appear in the result (the change really happened)
    --expect-gone T       T must no longer appear in the result
    --require-change      FAIL if nothing changed at all

Exit codes: 0 = only allowed changes, 1 = violations, 2 = usage error
"""
import argparse
import datetime
import hashlib
import json
import os
import signal
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from docgen.compare import check, kind_of  # noqa: E402

signal.signal(signal.SIGPIPE, signal.SIG_DFL)  # output piped into head: no traceback


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("original")
    ap.add_argument("result")
    for flag in ("text", "section", "style", "sheet", "cell", "feature", "part"):
        ap.add_argument(f"--allow-{flag}", action="append", default=[])
    for flag in ("slide", "notes", "page"):
        ap.add_argument(f"--allow-{flag}", action="append", type=int, default=[])
    for flag in ("band", "structure", "recalc", "theme"):
        ap.add_argument(f"--allow-{flag}", action="store_true")
    ap.add_argument("--expect-text", action="append", default=[])
    ap.add_argument("--expect-gone", action="append", default=[])
    ap.add_argument("--require-change", action="store_true")
    ap.add_argument("--text-only", action="store_true")
    ap.add_argument("--format-only", action="store_true")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--record", help="append a summary to this manifest JSON file")
    args = ap.parse_args()

    for p in (args.original, args.result):
        if not Path(p).exists():
            print(f"file not found: {p}", file=sys.stderr)
            return 2
    if not kind_of(args.original) or kind_of(args.original) != kind_of(args.result):
        print("usage error: both files must be the same supported format", file=sys.stderr)
        return 2

    scope = {"text": args.allow_text, "section": args.allow_section, "style": args.allow_style,
             "slide": args.allow_slide, "notes": args.allow_notes, "page": args.allow_page,
             "band": args.allow_band, "structure": args.allow_structure, "recalc": args.allow_recalc,
             "sheet": args.allow_sheet, "cell": args.allow_cell, "feature": args.allow_feature,
             "theme": args.allow_theme, "part": args.allow_part, "expect": args.expect_text,
             "gone": args.expect_gone, "require_change": args.require_change,
             "text_only": args.text_only, "format_only": args.format_only}
    rep = check(args.original, args.result, scope)

    if args.record:
        rec = {"time": datetime.datetime.now().isoformat(timespec="seconds"),
               "original": os.path.basename(args.original), "original_sha256": sha256(args.original),
               "result": os.path.basename(args.result), "result_sha256": sha256(args.result),
               "ok": rep["ok"], "changes": [f"{e['where']}: {e['what']}" for e in rep["allowed"]][:50],
               "changed_parts": (rep["parts"] or {}).get("semantic"), "scope": {k: v for k, v in scope.items() if v}}
        path = Path(args.record)
        data = json.loads(path.read_text()) if path.exists() else []
        data.append(rec)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2))

    if args.json:
        print(json.dumps(rep, ensure_ascii=False, indent=2, default=str))
    else:
        print(f"[diff_check] {'PASS' if rep['ok'] else 'FAIL'}  {args.result}  ({rep['kind']})")
        for e in rep["allowed"]:
            print(f"  ok   {e['where']}: {e['what']}  [{e['allowed_by']}]")
            if e.get("text_before") != e.get("text_after"):
                print(f"       before: {e.get('text_before', '')[:100]!r}\n       after:  {e.get('text_after', '')[:100]!r}")
        for e in rep["violations"]:
            print(f"  OUT OF SCOPE  {e['kind']} {e['where']}: {e['what']}")
            print(f"       before: {e.get('text_before', '')[:100]!r}\n       after:  {e.get('text_after', '')[:100]!r}")
        for n in rep["part_violations"]:
            print(f"  OUT OF SCOPE  package part changed: {n}")
        for m in rep["expectation_failures"]:
            print(f"  EXPECTATION FAILED  {m}")
        if rep["parts"]:
            p = rep["parts"]
            print(f"  parts: changed={p['semantic']} added={p['added']} removed={p['removed']}")
        if not rep["allowed"] and not rep["violations"]:
            print("  no content changes")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
