#!/usr/bin/env python3
"""DOCGEN read-back: print the content of an office or PDF file.

Usage:
    readback.py <file> [--max-chars N] [--tables] [--format]

- .docx/.xlsx/.pptx (+ legacy .doc/.xls/.ppt, templates .dotx/.potx/.xltx):
  office_oxide -> markdown
- .pdf: pdfplumber per-page text (and table summaries with --tables)
- --format: show formatting too (DOCX/PPTX runs: bold/italic/underline/
  color/size/font, paragraph alignment; PPTX shape position/size; XLSX cell
  formats; PDF line font/size/color/position) — use it to confirm that a
  style change landed exactly where it should.
"""
import argparse
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))

OFFICE_EXTS = {".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"}
TEMPLATE_EXTS = {".dotx": ".docx", ".docm": ".docx", ".potx": ".pptx", ".pptm": ".pptx",
                 ".xltx": ".xlsx", ".xlsm": ".xlsx"}


def as_plain_document(path: Path) -> Path:
    """Templates/macro files: same package with the normal main content type."""
    from docgen.package import Package
    fd, tmp = tempfile.mkstemp(suffix=TEMPLATE_EXTS[path.suffix.lower()])
    os.close(fd)
    Package(path).as_document().save(tmp)
    return Path(tmp)


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


# ---- --format -----------------------------------------------------------
def _docx_props(rpr, w):
    if rpr is None:
        return ""
    out = []
    for tag, label in (("b", "bold"), ("i", "italic"), ("strike", "strike")):
        el = rpr.find(w(tag))
        if el is not None:
            out.append(label if el.get(w("val"), "1") not in ("0", "false") else f"no-{label}")
    for tag, label, fmt in (("u", "underline", "{}"), ("color", "color", "#{}"), ("highlight", "highlight", "{}")):
        el = rpr.find(w(tag))
        if el is not None:
            out.append(f"{label}={fmt.format(el.get(w('val')))}")
    sz = rpr.find(w("sz"))
    if sz is not None:
        out.append(f"{int(sz.get(w('val'))) / 2:g}pt")
    f = rpr.find(w("rFonts"))
    if f is not None and f.get(w("ascii")):
        out.append(f"font={f.get(w('ascii'))}")
    return ",".join(out)


def format_docx(path):
    from docgen import docx as D
    from docgen.package import Package
    pkg = Package(path)
    lines = []
    for p in D.paragraphs(pkg):
        ppr = p.el.find(D.w("pPr"))
        jc = ppr.find(D.w("jc")) if ppr is not None else None
        sp = ppr.find(D.w("spacing")) if ppr is not None else None
        extra = f" align={jc.get(D.w('val'))}" if jc is not None else ""
        if sp is not None and sp.get(D.w("line")):
            extra += f" line={int(sp.get(D.w('line'))) / 240:g}"
        lines.append(f"[{p.where()} | style={p.style}{extra}] {p.text}")
        for r in p.runs:
            props = _docx_props(r.el.find(D.w("rPr")), D.w)
            if r.text and props:
                lines.append(f"    {r.text!r} {{{props}}}")
    for sid, el in D.Styles(pkg).by_id.items():
        if D.Styles(pkg).level(sid) is not None:
            props = _docx_props(el.find(D.w("rPr")), D.w)
            if props:
                lines.append(f"(style {sid}: {props})")
    return "\n".join(lines)


def format_pptx(path):
    from docgen import pptx as P
    from docgen.package import Package
    pkg = Package(path)
    out = []
    cm = lambda v: f"{v / 360000:.2f}"
    for sl in P.slides(pkg):
        out.append(f"=== slide {sl.index}{' (hidden)' if sl.hidden else ''}: {sl.title}")
        for sh in P.shapes(sl):
            geo = sh.xfrm
            pos = f" @({cm(geo[0])},{cm(geo[1])}) {cm(geo[2])}x{cm(geo[3])}cm" if geo else ""
            fill = sh.el.find(P.p("spPr") + "/" + P.a("solidFill") + "/" + P.a("srgbClr"))
            out.append(f"  [{sh.kind} {sh.name!r}{pos}{' fill=#' + fill.get('val') if fill is not None else ''}] "
                       f"{sh.text[:80]!r}")
            for para in P.paragraphs_of(sh):
                ppr = para.find(P.a("pPr"))
                for r in para.iter(P.a("r")):
                    rpr = r.find(P.a("rPr"))
                    if rpr is None:
                        continue
                    props = [f"{k}={v}" for k, v in rpr.attrib.items() if k in ("b", "i", "u", "strike")]
                    if rpr.get("sz"):
                        props.append(f"{int(rpr.get('sz')) / 100:g}pt")
                    clr = rpr.find(P.a("solidFill") + "/" + P.a("srgbClr"))
                    if clr is not None:
                        props.append(f"color=#{clr.get('val')}")
                    lat = rpr.find(P.a("latin"))
                    if lat is not None:
                        props.append(f"font={lat.get('typeface')}")
                    if ppr is not None and ppr.get("algn"):
                        props.append(f"align={ppr.get('algn')}")
                    if props:
                        out.append(f"      {r.findtext(P.a('t'))!r} {{{','.join(props)}}}")
        notes = P.get_notes(pkg, sl.index)
        if notes:
            out.append(f"  notes: {notes!r}")
    return "\n".join(out)


def format_xlsx(path):
    from docgen import xlsx as X
    from docgen.package import Package
    pkg = Package(path)
    st = pkg.xml("xl/styles.xml")
    S = X.s
    xfs, fonts, fills = (list(st.find(S(t))) for t in ("cellXfs", "fonts", "fills"))
    nf = st.find(S("numFmts"))
    numfmts = {n.get("numFmtId"): n.get("formatCode") for n in (nf if nf is not None else [])}
    out = []
    for name, part, state in X.sheets(pkg):
        out.append(f"=== sheet {name}{' (' + state + ')' if state != 'visible' else ''}")
        for c in pkg.xml(part).iter(S("c")):
            v, f = X.get(pkg, f"{name}!{c.get('r')}")
            xf = xfs[int(c.get("s", "0"))]
            props = []
            font = fonts[int(xf.get("fontId", "0"))]
            for tag in ("b", "i", "u", "strike"):
                if font.find(S(tag)) is not None:
                    props.append(tag)
            col = font.find(S("color"))
            if col is not None and col.get("rgb"):
                props.append(f"color=#{col.get('rgb')[-6:]}")
            for tag, key in (("sz", "size"), ("name", "font")):
                el = font.find(S(tag))
                if el is not None and tag == "name" and el.get("val") not in ("Calibri",):
                    props.append(f"{key}={el.get('val')}")
            fg = fills[int(xf.get("fillId", "0"))].find(S("patternFill") + "/" + S("fgColor"))
            if fg is not None and fg.get("rgb"):
                props.append(f"fill=#{fg.get('rgb')[-6:]}")
            if xf.get("numFmtId", "0") != "0":
                props.append(f"numfmt={numfmts.get(xf.get('numFmtId'), xf.get('numFmtId'))}")
            al = xf.find(S("alignment"))
            if al is not None and al.get("horizontal"):
                props.append(f"align={al.get('horizontal')}")
            if v is None and f is None and not props:
                continue
            out.append(f"  {c.get('r')}: {v!r}{' ' + f if f else ''}{' {' + ','.join(props) + '}' if props else ''}")
    return "\n".join(out)


def format_pdf(path):
    from docgen import pdf as F
    out = []
    for pno, ls in enumerate(F.lines(str(path)), 1):
        out.append(f"--- page {pno} ---")
        for l in ls:
            out.append(f"  y={l.top:6.1f} x={l.x0:6.1f} {F.base_font(l.font)} {l.size:g}pt "
                       f"rgb{l.color}{' MIXED' if l.mixed else ''}{' [' + l.band + ']' if l.band else ''}: {l.text}")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--max-chars", type=int, default=20000,
                    help="truncate output (default 20000)")
    ap.add_argument("--tables", action="store_true",
                    help="include table summaries for PDFs")
    ap.add_argument("--format", action="store_true", help="include formatting details")
    args = ap.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"file not found: {path}", file=sys.stderr)
        return 2

    try:
        ext = path.suffix.lower()
        if ext in TEMPLATE_EXTS:
            path, ext = as_plain_document(path), TEMPLATE_EXTS[ext]
        if args.format and ext in (".docx", ".pptx", ".xlsx", ".pdf"):
            content = {".docx": format_docx, ".pptx": format_pptx, ".xlsx": format_xlsx,
                       ".pdf": format_pdf}[ext](path)
        elif ext == ".pdf":
            content = read_pdf(path, args.tables)
        elif ext in OFFICE_EXTS:
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
