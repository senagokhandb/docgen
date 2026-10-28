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
  - decks: no text drawn over other text (e.g. bullets added as separate
    shapes on one spot); --baseline ORIGINAL: what the original already has
    is reported, not failed
  - the file has text but the PDF has none -> FAIL (blank render)
  - --expect-pages N, --expect-text T (repeatable)
Formula errors are not checked here: xlsx_recalc.py reads computed values
directly and knows which errors are new.

Usage:
    render_check.py <file> [-o rendered.pdf] [--baseline ORIGINAL] [--expect-pages N] [--expect-text T]... [--json]

Supported: .docx .pptx .xlsx (+ .dotx/.docm .potx/.pptm .xltx/.xlsm)
Exit codes: 0 = pass, 1 = fail, 2 = usage error
"""
import argparse
import io
import json
import os
import sys
import time
from collections import defaultdict
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
    """(kind, has_text, visible slide numbers or None) read from the package itself."""
    pkg = Package(path)
    kind = pkg.kind
    if kind == "docx":
        has_text = any((t.text or "").strip() for t in pkg.xml("word/document.xml").iter(D.w("t")))
        return kind, has_text, None
    if kind == "pptx":
        visible = [s for s in P.slides(pkg) if not s.hidden]
        has_text = any((t.text or "").strip() for s in visible for t in s.el.iter(P.a("t")))
        return kind, has_text, [s.index for s in visible]
    if kind == "xlsx":
        # hidden sheets are not printed, so only visible ones count
        has_text = any(c.find(X.s("v")) is not None or c.get("t") == "inlineStr"
                       for _, part, state in X.sheets(pkg) if state == "visible"
                       for c in pkg.xml(part).iter(X.s("c")))
        return kind, has_text, None
    raise ValueError("not an office package (no word/document.xml, ppt/presentation.xml or xl/workbook.xml)")


def squash(s):
    return " ".join(s.split())


def text_lines(page):
    """The page's text as lines: chars in drawing order that stay on one
    baseline and move right."""
    lines, cur = [], []
    for c in page.chars:
        if not c.get("upright", True):
            continue
        if cur and (abs(c["top"] - cur[-1]["top"]) > 0.5 * max(c["size"], cur[-1]["size"])
                    or c["x0"] < cur[-1]["x0"] - 0.5 * c["size"]):
            lines.append(cur)
            cur = []
        cur.append(c)
    if cur:
        lines.append(cur)
    return lines


def covers(a, b):
    """Glyph a and glyph b share at least half of the smaller one's box."""
    if max(a["size"], b["size"]) > 1.5 * min(a["size"], b["size"]):
        return False  # a big decorative glyph behind smaller text
    if a["text"] == b["text"] and abs(a["x0"] - b["x0"]) < 0.25 * a["size"] and abs(a["top"] - b["top"]) < 0.25 * a["size"]:
        return False  # the same glyph again, slightly offset: shadow or outline
    w = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
    h = min(a["bottom"], b["bottom"]) - max(a["top"], b["top"])
    area = min((a["x1"] - a["x0"]) * (a["bottom"] - a["top"]), (b["x1"] - b["x0"]) * (b["bottom"] - b["top"]))
    return w > 0 and h > 0 and w * h >= 0.5 * area


def collisions(page, min_glyphs=3):
    """[(text, text)] for pairs of text lines drawn over each other: at least
    min_glyphs visible glyphs of the two lines cover each other."""
    lines = text_lines(page)
    glyphs = [(li, c) for li, line in enumerate(lines) for c in line if c["text"].strip()]
    if not glyphs:
        return []
    step = max(sorted(c["size"] for _, c in glyphs)[len(glyphs) // 2], 1.0)
    grid = defaultdict(list)
    for k, (_, c) in enumerate(glyphs):
        for gx in range(int(c["x0"] // step), int(c["x1"] // step) + 1):
            for gy in range(int(c["top"] // step), int(c["bottom"] // step) + 1):
                grid[gx, gy].append(k)
    hit, done = defaultdict(set), set()
    for ks in grid.values():
        for i, a in enumerate(ks):
            for b in ks[i + 1:]:
                if (a, b) in done:
                    continue
                done.add((a, b))
                (la, ca), (lb, cb) = glyphs[a], glyphs[b]
                if la != lb and covers(ca, cb):
                    hit[la, lb].update((a, b))

    def text(li):
        return squash("".join(c["text"] for c in lines[li]))
    return [(text(la), text(lb)) for (la, lb), ks in hit.items() if len(ks) >= min_glyphs]


def page_collisions(pdf):
    """{page number: [(text, text)]} of a rendered PDF."""
    with pdfplumber.open(io.BytesIO(pdf)) as doc:
        return {n: c for n, pg in enumerate(doc.pages, 1) if (c := collisions(pg))}


def short(s, n=50):
    return s if len(s) <= n else s[:n - 1] + "…"


def check_collisions(rep, pdf, visible, baseline, url, timeout):
    """Text drawn over other text on a rendered deck: an error per slide; a
    note when the baseline (the original of an edit) has the same overlap."""
    found = page_collisions(pdf)
    if not found:
        return
    known = set()
    if baseline:
        try:
            base_pdf, _ = render(Path(baseline), url, timeout)
            known = {frozenset(pair) for pairs in page_collisions(base_pdf).values() for pair in pairs}
        except Exception as e:
            rep["errors"].append(f"cannot check the baseline — {e}")
    rep["collisions"] = []
    for page, pairs in found.items():
        slide = visible[page - 1] if visible and page <= len(visible) else page
        new = [p for p in pairs if frozenset(p) not in known]
        rep["collisions"] += [{"slide": slide, "texts": list(p), "in_baseline": frozenset(p) in known} for p in pairs]
        a, b = (new or pairs)[0]
        what = f"slide {slide}: text drawn over other text ({len(new or pairs)} pair(s)), e.g. {short(a)!r} / {short(b)!r}"
        if new:
            rep["errors"].append(what)
        else:
            rep["notes"].append(what + " — the original has it too")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("-o", "--out", help="also write the rendered PDF here (for readback or a visual look)")
    ap.add_argument("--baseline", help="the original of an edited deck: text over text it already has is "
                                       "reported, not failed")
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
    if a.baseline and not Path(a.baseline).is_file():
        print(f"usage error: no such baseline file: {a.baseline}", file=sys.stderr)
        return 2
    rep = {"file": str(path), "kind": KINDS[path.suffix.lstrip(".").lower()], "ok": False, "pages": None,
           "expected_pages": a.expect_pages, "text_chars": None, "seconds": None, "attempts": 0,
           "pdf": None, "errors": [], "notes": []}
    errors = rep["errors"]

    try:
        kind, has_text, visible = source_facts(path)
        rep["kind"] = kind
    except Exception as e:  # broken zip / XML: still let LibreOffice try, it is the real-world test
        kind, has_text, visible = rep["kind"], False, None
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
    if visible is not None:
        rep["visible_slides"] = len(visible)
        if rep["pages"] != len(visible):
            errors.append(f"{rep['pages']} PDF pages but {len(visible)} visible slides")
    if kind == "pptx":
        check_collisions(rep, pdf, visible, a.baseline, url, a.timeout)
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
        for n in rep["notes"]:
            print(f"  note: {n}")
        for e in rep["errors"]:
            print(f"  ERROR: {e}")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
