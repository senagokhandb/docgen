#!/usr/bin/env python3
"""DOCGEN render check: let LibreOffice (through Gotenberg) open an office
file and lay it out as PDF, the way a user's office suite would, and FAIL
unless that really produced pages.

Use it instead of `curl -s .../libreoffice/convert -o x.pdf && ls x.pdf`:
without -f, curl saves Gotenberg's 4xx/5xx error text as "x.pdf" and exits 0,
so a failed render looks like a success.

Checks:
  - Gotenberg answers 200 and the body is a PDF that opens with >= 1 page
  - decks: PDF pages == visible slides (hidden slides are not exported)
  - the file has text but the PDF has none -> FAIL (blank render)
  - --expect-pages N, --expect-text T (repeatable)
Formula errors are not checked here: xlsx_recalc.py reads computed values
directly and knows which errors are new.

Usage:
    render_check.py <file> [-o rendered.pdf] [--expect-pages N] [--expect-text T]... [--json]

Supported: .docx .pptx .xlsx (+ .dotx/.docm .potx/.pptm .xltx/.xlsm)
Exit codes: 0 = pass, 1 = fail, 2 = usage error
"""
import argparse
import io
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
import pdfplumber  # noqa: E402
import requests  # noqa: E402

from docgen import docx as D  # noqa: E402
from docgen import pptx as P  # noqa: E402
from docgen import xlsx as X  # noqa: E402
from docgen.package import Package  # noqa: E402

KINDS = {"docx": "docx", "dotx": "docx", "docm": "docx", "pptx": "pptx", "potx": "pptx", "pptm": "pptx",
         "xlsx": "xlsx", "xltx": "xlsx", "xlsm": "xlsx"}
RETRY_STATUS = (502, 503, 504)  # Gotenberg busy or LibreOffice restarting


def render(path, url, timeout):
    """-> (pdf bytes, attempts). Raises RuntimeError with Gotenberg's own message."""
    last = None
    for attempt in (1, 2):
        try:
            with open(path, "rb") as fh:
                # an ASCII name with the real extension: Gotenberg picks the import filter by it
                resp = requests.post(f"{url}/forms/libreoffice/convert",
                                     files={"files": (f"document{path.suffix.lower()}", fh)}, timeout=timeout)
        except requests.RequestException as e:
            last = f"Gotenberg not reachable at {url}: {type(e).__name__}: {e}"
        else:
            if resp.status_code == 200:
                if not resp.content.startswith(b"%PDF"):
                    raise RuntimeError(f"Gotenberg answered 200 but the body is not a PDF: {resp.content[:120]!r}")
                return resp.content, attempt
            last = f"Gotenberg HTTP {resp.status_code}: {resp.text.strip()[:300]}"
            if resp.status_code not in RETRY_STATUS:
                break
        if attempt == 1:
            time.sleep(3)
    raise RuntimeError(last)


def source_facts(path):
    """(kind, has_text, visible_slides or None) read from the package itself."""
    pkg = Package(path)
    kind = pkg.kind
    if kind == "docx":
        has_text = any((t.text or "").strip() for t in pkg.xml("word/document.xml").iter(D.w("t")))
        return kind, has_text, None
    if kind == "pptx":
        visible = [s for s in P.slides(pkg) if not s.hidden]
        has_text = any((t.text or "").strip() for s in visible for t in s.el.iter(P.a("t")))
        return kind, has_text, len(visible)
    if kind == "xlsx":
        # hidden sheets are not printed, so only visible ones count
        has_text = any(c.find(X.s("v")) is not None or c.get("t") == "inlineStr"
                       for _, part, state in X.sheets(pkg) if state == "visible"
                       for c in pkg.xml(part).iter(X.s("c")))
        return kind, has_text, None
    raise ValueError("not an office package (no word/document.xml, ppt/presentation.xml or xl/workbook.xml)")


def squash(s):
    return " ".join(s.split())


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("-o", "--out", help="also write the rendered PDF here (for readback or a visual look)")
    ap.add_argument("--expect-pages", type=int, help="the PDF must have exactly N pages")
    ap.add_argument("--expect-text", action="append", default=[], help="text that must appear in the PDF (repeatable)")
    ap.add_argument("--timeout", type=int, default=200, help="seconds per Gotenberg request (default 200)")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    a = ap.parse_args()

    path = Path(a.file)
    if not path.is_file() or path.suffix.lstrip(".").lower() not in KINDS:
        print(f"usage error: pass an existing {'/'.join('.' + k for k in KINDS)} file", file=sys.stderr)
        return 2
    url = os.environ.get("GOTENBERG_URL", "http://gotenberg:3000").rstrip("/")
    rep = {"file": str(path), "kind": KINDS[path.suffix.lstrip(".").lower()], "ok": False, "pages": None,
           "expected_pages": a.expect_pages, "text_chars": None, "seconds": None, "attempts": 0,
           "pdf": None, "errors": []}
    errors = rep["errors"]

    try:
        kind, has_text, visible_slides = source_facts(path)
        rep["kind"] = kind
    except Exception as e:  # broken zip / XML: still let LibreOffice try, it is the real-world test
        kind, has_text, visible_slides = rep["kind"], False, None
        errors.append(f"cannot read the package: {type(e).__name__}: {e}")

    t0 = time.monotonic()
    try:
        pdf, rep["attempts"] = render(path, url, a.timeout)
    except RuntimeError as e:
        rep["seconds"] = round(time.monotonic() - t0, 1)
        errors.append(f"render failed — {e}")
        return emit(rep, a.json)
    rep["seconds"] = round(time.monotonic() - t0, 1)

    if a.out:
        Path(a.out).write_bytes(pdf)
        rep["pdf"] = a.out
    try:
        with pdfplumber.open(io.BytesIO(pdf)) as doc:
            rep["pages"] = len(doc.pages)
            text = "\n".join(pg.extract_text() or "" for pg in doc.pages)
    except Exception as e:
        errors.append(f"rendered PDF does not open: {type(e).__name__}: {e}")
        return emit(rep, a.json)
    rep["text_chars"] = len(text.strip())

    if rep["pages"] < 1:
        errors.append("rendered PDF has no pages")
    if visible_slides is not None:
        rep["visible_slides"] = visible_slides
        if rep["pages"] != visible_slides:
            errors.append(f"{rep['pages']} PDF pages but {visible_slides} visible slides")
    if a.expect_pages is not None and rep["pages"] != a.expect_pages:
        errors.append(f"{rep['pages']} PDF pages, expected {a.expect_pages}")
    if has_text and not rep["text_chars"]:
        errors.append("the file has text but the rendered PDF has none (blank render)")
    flat = squash(text)
    for t in a.expect_text:
        if squash(t) not in flat:
            errors.append(f"expected text not in the rendered PDF: {t!r}")
    return emit(rep, a.json)


def emit(rep, as_json):
    rep["ok"] = not rep["errors"]
    if as_json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        status = "PASS" if rep["ok"] else "FAIL"
        pages = f", {rep['pages']} pages" if rep["pages"] is not None else ""
        print(f"[render_check] {status}  {rep['file']}  ({rep['kind']}{pages}, {rep['seconds']}s)")
        if "visible_slides" in rep:
            print(f"  visible slides: {rep['visible_slides']}")
        if rep["text_chars"] is not None:
            print(f"  rendered text chars: {rep['text_chars']}")
        if rep["pdf"]:
            print(f"  pdf: {rep['pdf']}")
        for e in rep["errors"]:
            print(f"  ERROR: {e}")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
