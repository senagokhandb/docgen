"""Independent result checks for the DOCGEN eval.

Deliberately NOT built on lib/docgen: results are read with python-docx,
python-pptx, openpyxl, pdfplumber and raw OOXML so the system under test
never grades itself. Formula results are compared with a fresh LibreOffice
recalculation. Every check returns (ok: bool, observed: str).
"""
import colorsys
import difflib
import fnmatch
import hashlib
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def w(t):
    return f"{{{W}}}{t}"


def a(t):
    return f"{{{A}}}{t}"


def p(t):
    return f"{{{P}}}{t}"


def norm(t):
    return re.sub(r"\s+", " ", (t or "").replace(" ", " ")).strip()


# ---- colors -------------------------------------------------------------
FAMILIES = {
    "red": lambda h, s, v: (h < 0.04 or h > 0.95) and s > 0.45 and v > 0.35,
    "darkred": lambda h, s, v: (h < 0.04 or h > 0.95) and s > 0.45 and 0.2 < v <= 0.8,
    "green": lambda h, s, v: 0.2 < h < 0.48 and s > 0.3 and v > 0.2,
    "darkgreen": lambda h, s, v: 0.2 < h < 0.48 and s > 0.3 and 0.15 < v <= 0.65,
    "blue": lambda h, s, v: 0.52 < h < 0.72 and s > 0.3 and v > 0.2,
    "navy": lambda h, s, v: 0.52 < h < 0.72 and s > 0.3 and v <= 0.6,
    "yellow": lambda h, s, v: 0.11 < h < 0.2 and s > 0.4 and v > 0.6,
    "orange": lambda h, s, v: 0.04 <= h <= 0.11 and s > 0.5 and v > 0.6,
    "lightgray": lambda h, s, v: s < 0.12 and 0.7 <= v < 0.99,
    "gray": lambda h, s, v: s < 0.12 and 0.3 <= v < 0.99,
    "white": lambda h, s, v: s < 0.06 and v >= 0.97,
    "black": lambda h, s, v: v < 0.15,
    "purple": lambda h, s, v: 0.72 <= h <= 0.92 and s > 0.3,
    "lightblue": lambda h, s, v: 0.5 < h < 0.7 and s > 0.08 and v > 0.75,
}
HIGHLIGHT = {"yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF", "red": "FF0000", "blue": "0000FF",
             "darkYellow": "808000", "lightGray": "C0C0C0", "darkRed": "800000", "magenta": "FF00FF"}


def color_ok(hexval, want):
    """want: exact 'RRGGBB' or a family name."""
    if not hexval:
        return False
    hexval = hexval.upper()[-6:]
    if re.fullmatch(r"[0-9A-Fa-f]{6}", want or ""):
        return hexval == want.upper()
    try:
        r, g, b = (int(hexval[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return False
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return FAMILIES[want](h, s, v)


# ---- generic --------------------------------------------------------------
def parts(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist() if not i.is_dir()}, \
            [i.filename for i in z.infolist() if i.is_dir()]


def check_parts_only(ctx, allow=()):
    """Byte-level: only the allowed package parts may differ (docProps always)."""
    a_, _ = parts(ctx.original)
    b_, dirs = parts(ctx.output)
    allow = list(allow) + ["docProps/*"]
    match = lambda n: any(n == g or fnmatch.fnmatchcase(n, g.replace("[", "[[]")) for g in allow)
    bad = sorted(n for n in set(a_) | set(b_) if a_.get(n) != b_.get(n) and not match(n))
    if dirs:
        bad.append(f"zip dir entries {dirs[:2]}")
    return not bad, f"unexpected part changes: {bad[:6]}" if bad else "only allowed parts changed"


def check_identical(ctx):
    same = hashlib.sha256(Path(ctx.original).read_bytes()).hexdigest() == \
        hashlib.sha256(Path(ctx.output).read_bytes()).hexdigest()
    return same, "byte-identical" if same else "file differs from the upload"


def check_opens(ctx):
    try:
        ext = Path(ctx.output).suffix.lower()
        if ext == ".docx":
            import docx
            docx.Document(ctx.output)
        elif ext == ".pptx":
            import pptx
            pptx.Presentation(ctx.output)
        elif ext == ".xlsx":
            import openpyxl
            openpyxl.load_workbook(ctx.output)
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(ctx.output) as pdf:
                assert len(pdf.pages) > 0
        return True, "opens"
    except Exception as e:  # noqa: BLE001
        return False, f"cannot open: {e}"


def all_text(path):
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return "\n".join(norm("".join(t.text or "" for t in p_.iter(w("t")))) for p_ in docx_paragraph_elements(path))
    if ext == ".pptx":
        import pptx
        prs = pptx.Presentation(path)
        out = []
        for s in prs.slides:
            for sh in s.shapes:
                if sh.has_text_frame:
                    out.append(sh.text_frame.text)
                if getattr(sh, "has_table", False) and sh.has_table:
                    out += [c.text for r in sh.table.rows for c in r.cells]
            if s.has_notes_slide:
                out.append(s.notes_slide.notes_text_frame.text)
        return "\n".join(out)
    if ext == ".xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(path)
        return "\n".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value is not None)
    if ext == ".pdf":
        return "\n".join(pdf_page_text(path))
    return ""


def check_text_has(ctx, texts):
    t = norm(all_text(ctx.output))
    miss = [x for x in texts if norm(x) not in t]
    return not miss, f"missing {miss}" if miss else "all present"


def check_text_not(ctx, texts):
    t = norm(all_text(ctx.output))
    found = [x for x in texts if norm(x) in t]
    return not found, f"still present {found}" if found else "all absent"


def check_text_count(ctx, text, n, whole_word=True):
    t = all_text(ctx.output)
    pat = rf"(?<!\w){re.escape(text)}(?!\w)" if whole_word else re.escape(text)
    k = len(re.findall(pat, t))
    return k == n, f"{text!r} occurs {k}x (expected {n})"


# ---- DOCX -----------------------------------------------------------------
def _xml(path, part):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(part)) if part in z.namelist() else None


def _in_fallback(el):
    return any(x.tag in (f"{{{MC}}}Fallback", w("del")) for x in el.iterancestors())


def docx_paragraph_elements(path, parts_=("word/document.xml",), headers=True):
    out = []
    with zipfile.ZipFile(path) as z:
        names = list(parts_) + ([n for n in z.namelist() if re.match(r"word/(header|footer)\d*\.xml$", n)] if headers else [])
        for n in names:
            if n in z.namelist():
                root = etree.fromstring(z.read(n))
                out += [x for x in root.iter(w("p")) if not _in_fallback(x)]
    return out


class DocxStyles:
    def __init__(self, path):
        self.root = _xml(path, "word/styles.xml")
        self.by_id = {s.get(w("styleId")): s for s in self.root.iter(w("style"))} if self.root is not None else {}
        self.default_p = next((i for i, s in self.by_id.items() if s.get(w("type")) == "paragraph"
                               and s.get(w("default")) in ("1", "true")), None)
        dd = self.root.find(f"{w('docDefaults')}/{w('rPrDefault')}/{w('rPr')}") if self.root is not None else None
        self.defaults = dd
        th = _xml(path, "word/theme/theme1.xml")
        self.theme_fonts = {}
        if th is not None:
            for k in ("major", "minor"):
                lat = th.find(f".//{{{A}}}{k}Font/{{{A}}}latin")
                if lat is not None:
                    self.theme_fonts[k] = lat.get("typeface")

    def chain(self, sid):
        out = []
        while sid and sid in self.by_id and sid not in out:
            out.append(sid)
            b = self.by_id[sid].find(w("basedOn"))
            sid = b.get(w("val")) if b is not None else None
        return out

    def name(self, sid):
        n = self.by_id.get(sid)
        n = n.find(w("name")) if n is not None else None
        return (n.get(w("val")) if n is not None else sid or "").lower()

    def level(self, sid):
        for s in self.chain(sid):
            nm = self.name(s)
            if nm == "title":
                return 0
            m = re.fullmatch(r"heading (\d)", nm)
            if m:
                return int(m.group(1))
            ol = self.by_id[s].find(f"{w('pPr')}/{w('outlineLvl')}")
            if ol is not None and ol.get(w("val")) != "9":
                return int(ol.get(w("val"))) + 1
        return None

    def rpr_layers(self, para, run):
        """rPr elements from most to least specific."""
        layers = [run.find(w("rPr"))]
        rs = run.find(f"{w('rPr')}/{w('rStyle')}")
        if rs is not None:
            layers += [self.by_id[s].find(w("rPr")) for s in self.chain(rs.get(w("val")))]
        ps = para.find(f"{w('pPr')}/{w('pStyle')}")
        sid = ps.get(w("val")) if ps is not None else self.default_p
        layers += [self.by_id[s].find(w("rPr")) for s in self.chain(sid)]
        layers.append(self.defaults)
        return [x for x in layers if x is not None]

    def ppr_layers(self, para):
        layers = [para.find(w("pPr"))]
        ps = para.find(f"{w('pPr')}/{w('pStyle')}")
        sid = ps.get(w("val")) if ps is not None else self.default_p
        layers += [self.by_id[s].find(w("pPr")) for s in self.chain(sid)]
        return [x for x in layers if x is not None]

    def para_style(self, para):
        ps = para.find(f"{w('pPr')}/{w('pStyle')}")
        return ps.get(w("val")) if ps is not None else self.default_p


def _onoff(el):
    return None if el is None else el.get(w("val"), "true") not in ("0", "false", "none")


def docx_effective(styles, para, run):
    """Effective run properties (direct > char style > para style > defaults)."""
    out = {}
    for prop, tag in (("bold", "b"), ("italic", "i"), ("strike", "strike")):
        out[prop] = next((_onoff(l.find(w(tag))) for l in styles.rpr_layers(para, run) if l.find(w(tag)) is not None), False)
    u = next((l.find(w("u")) for l in styles.rpr_layers(para, run) if l.find(w("u")) is not None), None)
    out["underline"] = u is not None and u.get(w("val"), "single") != "none"
    c = next((l.find(w("color")) for l in styles.rpr_layers(para, run) if l.find(w("color")) is not None), None)
    out["color"] = (c.get(w("val")) or "").upper() if c is not None and (c.get(w("val")) or "").lower() != "auto" else None
    sz = next((l.find(w("sz")) for l in styles.rpr_layers(para, run) if l.find(w("sz")) is not None), None)
    out["size"] = int(sz.get(w("val"))) / 2 if sz is not None else None
    f = next((l.find(w("rFonts")) for l in styles.rpr_layers(para, run) if l.find(w("rFonts")) is not None), None)
    font = None
    if f is not None:
        font = f.get(w("ascii")) or f.get(w("hAnsi"))
        if not font and f.get(w("asciiTheme")):
            font = styles.theme_fonts.get("major" if "major" in f.get(w("asciiTheme")) else "minor")
    out["font"] = font
    hl = run.find(f"{w('rPr')}/{w('highlight')}")
    shd = run.find(f"{w('rPr')}/{w('shd')}")
    out["highlight"] = hl.get(w("val")) if hl is not None else (
        (shd.get(w("fill")) or "").upper() if shd is not None and shd.get(w("fill")) not in (None, "auto") else None)
    return out


def docx_chars(path):
    """[(para_el, text, [(char, run_el)])] for every paragraph."""
    res = []
    for para in docx_paragraph_elements(path):
        chars = []
        for r in para.iter(w("r")):
            if _in_fallback(r):
                continue
            for t in r.iter(w("t")):
                chars += [(ch, r) for ch in (t.text or "")]
        res.append((para, "".join(c for c, _ in chars), chars))
    return res


def _prop_ok(name, got, want):
    if name == "color":
        return color_ok(got, want)
    if name == "highlight":
        if got is None:
            return False
        hexv = HIGHLIGHT.get(got, got)
        return got == want or color_ok(hexv, want) or (want in HIGHLIGHT and hexv == HIGHLIGHT[want])
    if name == "size":
        return got is not None and abs(float(got) - float(want)) < 0.26
    if name == "font":
        return (got or "").lower() == str(want).lower()
    return bool(got) == bool(want)


def _find_span(text, target, near=None, occurrence=None):
    hits = [m.start() for m in re.finditer(re.escape(target), text)]
    if near:
        hits = [h for h in hits if near in text]
    if occurrence:
        hits = hits[occurrence - 1:occurrence]
    return hits


def check_docx_chars(ctx, text, props, near=None, occurrence=None, keep=None):
    """Every character of ``text`` has ``props``; ``keep`` = props that must
    equal the original's for the same characters."""
    st = DocxStyles(ctx.output)
    found = 0
    for para, ptext, chars in docx_chars(ctx.output):
        if near and near not in ptext:
            continue
        for start in _find_span(ptext, text, None, None):
            found += 1
            for i in range(start, start + len(text)):
                ch, run = chars[i]
                if not ch.strip():
                    continue
                eff = docx_effective(st, para, run)
                for k, v in props.items():
                    if not _prop_ok(k, eff.get(k), v):
                        return False, f"{text!r}: char {ch!r} {k}={eff.get(k)!r}, expected {v!r}"
    if not found:
        return False, f"{text!r} not found"
    if keep:
        ok, msg = _docx_keep(ctx, text, keep, near)
        if not ok:
            return ok, msg
    return True, f"{text!r} has {props}" + (f", kept {keep}" if keep else "")


def _docx_keep(ctx, text, keep, near=None):
    def props_of(path):
        st = DocxStyles(path)
        for para, ptext, chars in docx_chars(path):
            if near and near not in ptext:
                continue
            i = ptext.find(text)
            if i >= 0:
                return [docx_effective(st, para, chars[j][1]) for j in range(i, i + len(text)) if chars[j][0].strip()]
        return None
    a_, b_ = props_of(ctx.original), props_of(ctx.output)
    if a_ is None or b_ is None:
        return False, f"{text!r} not found for keep-check"
    for x, y in zip(a_, b_):
        for k in keep:
            if x.get(k) != y.get(k):
                return False, f"{text!r}: {k} changed {x.get(k)!r} -> {y.get(k)!r}"
    return True, "kept"


def check_docx_para(ctx, contains, align=None, line_spacing=None, style_level=None, space_after=None):
    st = DocxStyles(ctx.output)
    paras = [(pa, t) for pa, t, _ in docx_chars(ctx.output) if contains in t]
    if not paras:
        return False, f"paragraph with {contains!r} not found"
    for pa, t in paras:
        layers = st.ppr_layers(pa)
        got = []
        if align:
            jc = next((l.find(w("jc")).get(w("val")) for l in layers if l.find(w("jc")) is not None), "left")
            jc = {"start": "left", "end": "right", "distribute": "both"}.get(jc, jc)
            if jc != {"justify": "both"}.get(align, align):
                return False, f"{contains!r}: align={jc}, expected {align}"
            got.append(f"align={jc}")
        if line_spacing:
            sp = next((l.find(w("spacing")) for l in layers if l.find(w("spacing")) is not None
                       and l.find(w("spacing")).get(w("line"))), None)
            val = int(sp.get(w("line"))) / 240 if sp is not None else 1.0
            if abs(val - line_spacing) > 0.05:
                return False, f"{contains!r}: line spacing {val}, expected {line_spacing}"
            got.append(f"line spacing {val}")
        if space_after is not None:
            sp = next((l.find(w("spacing")) for l in layers if l.find(w("spacing")) is not None
                       and l.find(w("spacing")).get(w("after"))), None)
            val = int(sp.get(w("after"))) / 20 if sp is not None else 0
            if abs(val - space_after) > 0.6:
                return False, f"{contains!r}: space after {val}pt, expected {space_after}"
            got.append(f"space after {val}pt")
    return True, f"{contains!r}: " + (", ".join(got) or "paragraph properties ok")


def check_docx_headings(ctx, levels, props, others_unchanged=True):
    """Every heading paragraph of the given levels shows ``props``; with
    others_unchanged, non-heading paragraphs keep their effective color/size/font."""
    st = DocxStyles(ctx.output)
    n = 0
    for para, t, chars in docx_chars(ctx.output):
        lvl = st.level(st.para_style(para))
        if lvl in levels and t.strip():
            n += 1
            for ch, run in chars:
                if not ch.strip():
                    continue
                eff = docx_effective(st, para, run)
                for k, v in props.items():
                    if not _prop_ok(k, eff.get(k), v):
                        return False, f"heading {t!r}: {k}={eff.get(k)!r}, expected {v!r}"
    if not n:
        return False, "no heading paragraphs of those levels"
    if others_unchanged:
        ok, msg = _docx_body_same(ctx, levels)
        if not ok:
            return ok, msg
    return True, f"{n} headings ok"


def _docx_body_same(ctx, levels):
    def body(path):
        st = DocxStyles(path)
        out = []
        for para, t, chars in docx_chars(path):
            if st.level(st.para_style(para)) in levels or not t.strip():
                continue
            out.append((t, [tuple(sorted(docx_effective(st, para, r).items())) for ch, r in chars if ch.strip()][:3]))
        return out
    a_, b_ = body(ctx.original), body(ctx.output)
    if a_ != b_:
        diff = next(((x, y) for x, y in zip(a_, b_) if x != y), (len(a_), len(b_)))
        return False, f"non-heading text/format changed: {str(diff)[:160]}"
    return True, "body unchanged"


def _docx_sig(path):
    st = DocxStyles(path)
    out = []
    for para, t, chars in docx_chars(path):
        fmt = []
        for ch, r in chars:
            rp = r.find(w("rPr"))
            fmt.append(etree.tostring(rp, method="c14n") if rp is not None else b"")
        pp = para.find(w("pPr"))
        out.append((t, tuple(fmt), etree.tostring(pp, method="c14n") if pp is not None else b"", st.para_style(para)))
    return out


def check_docx_others_same(ctx, allow=(), allow_insert=(), allow_delete=()):
    """Paragraphs (body, tables, headers/footers) not mentioning an allowed
    text keep exactly their text, character and paragraph formatting."""
    a_, b_ = _docx_sig(ctx.original), _docx_sig(ctx.output)
    ok_txt = lambda t, extra=(): any(x and x in t for x in tuple(allow) + tuple(extra))
    sm = difflib.SequenceMatcher(a=[x[0] for x in a_], b=[x[0] for x in b_], autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            for x, y in zip(a_[i1:i2], b_[j1:j2]):
                if x != y and not ok_txt(x[0]):
                    return False, f"paragraph {x[0][:50]!r} formatting changed"
            continue
        if op == "replace" and i2 - i1 == j2 - j1:
            for x, y in zip(a_[i1:i2], b_[j1:j2]):
                if x != y and not (ok_txt(x[0]) or ok_txt(y[0])):
                    return False, f"paragraph changed: {x[0][:45]!r} -> {y[0][:45]!r}"
            continue
        for x in a_[i1:i2]:
            if x[0].strip() and not ok_txt(x[0], allow_delete):
                return False, f"paragraph removed: {x[0][:60]!r}"
        for y in b_[j1:j2]:
            if y[0].strip() and not ok_txt(y[0], allow_insert):
                return False, f"paragraph added: {y[0][:60]!r}"
    return True, "other paragraphs unchanged"


def _docx_tables(path):
    root = _xml(path, "word/document.xml")
    body = root.find(w("body"))
    out = []
    for t in body.iter(w("tbl")):
        rows = []
        for tr in t.findall(w("tr")):
            rows.append([norm("".join(x.text or "" for x in tc.iter(w("t")))) for tc in tr.findall(w("tc"))])
        out.append((t, rows))
    return out


def check_docx_table_cell(ctx, table_has, row, col, value):
    for t, rows in _docx_tables(ctx.output):
        if any(table_has in c for r in rows for c in r):
            hdr = rows[0]
            ci = next((i for i, h in enumerate(hdr) if col in h), None)
            rr = next((r for r in rows if r and r[0] == row), None)
            if ci is None or rr is None:
                return False, f"row {row!r}/col {col!r} not found in {hdr}"
            return rr[ci] == value, f"{row}/{col} = {rr[ci]!r} (expected {value!r})"
    return False, f"table with {table_has!r} not found"


def check_docx_table_shape(ctx, table_has, cols=None, rows=None, header=None, row_values=None):
    for t, rws in _docx_tables(ctx.output):
        if any(table_has in c for r in rws for c in r):
            if cols is not None and any(len(r) != cols for r in rws):
                return False, f"column counts {[len(r) for r in rws]}, expected {cols}"
            if rows is not None and len(rws) != rows:
                return False, f"{len(rws)} rows, expected {rows}"
            if header and header not in rws[0]:
                return False, f"header {rws[0]} lacks {header!r}"
            if row_values and row_values not in rws:
                return False, f"row {row_values} not in table"
            return True, f"table {len(rws)}x{len(rws[0])} ok"
    return False, f"table with {table_has!r} not found"


def check_docx_cell_fill(ctx, table_has, row, color):
    for t, rws in _docx_tables(ctx.output):
        if any(table_has in c for r in rws for c in r):
            tr = t.findall(w("tr"))[row]
            for tc in tr.findall(w("tc")):
                shd = tc.find(f"{w('tcPr')}/{w('shd')}")
                if shd is None or not color_ok(shd.get(w("fill")), color):
                    return False, f"cell fill {None if shd is None else shd.get(w('fill'))}, expected {color}"
            return True, "row fill ok"
    return False, "table not found"


def check_docx_order(ctx, first, then):
    """Paragraph containing ``then`` comes right after (the section of) ``first``."""
    texts = [t for _, t, _ in docx_chars(ctx.output)]
    i = next((k for k, t in enumerate(texts) if first in t), None)
    j = next((k for k, t in enumerate(texts) if then in t), None)
    if i is None or j is None:
        return False, f"missing {first!r} or {then!r}"
    return i < j, f"{first!r}@{i} before {then!r}@{j}"


def check_docx_before(ctx, text, next_heading):
    texts = [t for _, t, _ in docx_chars(ctx.output)]
    i = next((k for k, t in enumerate(texts) if text in t), None)
    j = next((k for k, t in enumerate(texts) if t.strip() == next_heading), None)
    if i is None or j is None:
        return False, f"missing {text!r} or heading {next_heading!r}"
    return i == j - 1 or (i < j and all(not texts[k].strip() for k in range(i + 1, j))), \
        f"new paragraph at {i}, next heading at {j}"


def check_docx_media(ctx, count=None, header_image=False):
    with zipfile.ZipFile(ctx.output) as z:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
        if count is not None and len(media) != count:
            return False, f"{len(media)} media parts, expected {count}"
        body = z.read("word/document.xml")
        drawings = body.count(b"<w:drawing")
        if count is not None and drawings < count:
            return False, f"{drawings} drawings in body, expected {count}"
        if header_image:
            hdrs = [z.read(n) for n in z.namelist() if re.match(r"word/header\d*\.xml$", n)]
            if not any(b"drawing" in h or b"pict" in h for h in hdrs) or not media:
                return False, "header image missing"
    return True, "media ok"


def check_docx_heading_texts(ctx, present=(), absent=(), min_count=0):
    st = DocxStyles(ctx.output)
    heads = [t.strip() for para, t, _ in docx_chars(ctx.output) if st.level(st.para_style(para)) is not None and t.strip()]
    miss = [h for h in present if not any(h.lower() in x.lower() for x in heads)]
    bad = [h for h in absent if any(h == x for x in heads)]
    if miss or bad or len(heads) < min_count:
        return False, f"headings {heads}; missing {miss}; should be gone {bad}"
    return True, f"headings {heads}"


def check_docx_style_used(ctx, min_heading_paras=1):
    """Headings use real heading styles (template styles), not manual formatting."""
    st = DocxStyles(ctx.output)
    n = sum(1 for para, t, _ in docx_chars(ctx.output) if t.strip() and st.level(st.para_style(para)) is not None)
    return n >= min_heading_paras, f"{n} paragraphs with heading styles"


def check_docx_numbered(ctx, min_items=2):
    root = _xml(ctx.output, "word/document.xml")
    n = len(root.findall(f".//{w('numPr')}"))
    st = DocxStyles(ctx.output)
    n += sum(1 for para, t, _ in docx_chars(ctx.output)
             if any(l.find(w("numPr")) is not None for l in st.ppr_layers(para)) and para.find(f".//{w('numPr')}") is None)
    return n >= min_items, f"{n} numbered paragraphs"


def check_docx_theme(ctx, slot=None, color=None, major=None, minor=None):
    th = _xml(ctx.output, "word/theme/theme1.xml")
    if th is None:
        return False, "no theme"
    if slot:
        el = th.find(f".//{{{A}}}clrScheme/{{{A}}}{slot}")
        v = el[0].get("val") if el is not None and len(el) else None
        if not color_ok(v, color):
            return False, f"theme {slot}={v}, expected {color}"
    for k, want in (("major", major), ("minor", minor)):
        if want:
            lat = th.find(f".//{{{A}}}{k}Font/{{{A}}}latin")
            if lat is None or lat.get("typeface") != want:
                return False, f"theme {k} font {None if lat is None else lat.get('typeface')}"
    return True, "theme ok"


def _docx_used_styles(z):
    """Style ids that reach the rendered document: referenced from the body,
    headers, footers, notes or numbering, their basedOn/link chains and the
    default styles."""
    used = set()
    for n in z.namelist():
        if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|numbering)\.xml$", n):
            used |= {e.get(w("val")) for e in etree.fromstring(z.read(n)).iter(w("pStyle"), w("rStyle"), w("tblStyle"))}
    styles = {s.get(w("styleId")): s for s in etree.fromstring(z.read("word/styles.xml")).iter(w("style"))} \
        if "word/styles.xml" in z.namelist() else {}
    used |= {sid for sid, s in styles.items() if s.get(w("default")) in ("1", "true")}
    todo = list(used)
    while todo:
        s = styles.get(todo.pop())
        for ref in (s.find(w("basedOn")), s.find(w("link"))) if s is not None else ():
            if ref is not None and ref.get(w("val")) not in used:
                used.add(ref.get(w("val")))
                todo.append(ref.get(w("val")))
    return used


def check_docx_theme_cache(ctx, tol=16):
    """Every cached hex value of a theme-bound color (w:color/@w:val, w:shd/@w:fill,
    borders, underline) that reaches the rendered document matches the theme
    as Word computes it ([MS-OI29500] 2.1.72: HSL lightness, shade L*s, tint
    L*t + 1-t, tint wins). LibreOffice, Gotenberg and PDF exports draw these
    caches, so a stale one shows the old color outside Word. Unused styles
    are not counted; ``tol`` (per channel) accepts a tint computed slightly
    differently but not a value left from the old theme."""
    th = _xml(ctx.output, "word/theme/theme1.xml")
    cs = th.find(f".//{{{A}}}clrScheme") if th is not None else None
    if cs is None:
        return False, "no theme"
    base = {}
    for s in cs:
        c = s[0] if len(s) else None
        v = c.get("lastClr") if c is not None and etree.QName(c).localname == "sysClr" else (c.get("val") if c is not None else None)
        base[etree.QName(s).localname] = v
    mapping = {"t1": "dark1", "bg1": "light1", "t2": "dark2", "bg2": "light2"}
    st = _xml(ctx.output, "word/settings.xml")
    m = st.find(w("clrSchemeMapping")) if st is not None else None
    mapping.update({etree.QName(k).localname: v for k, v in (m.attrib.items() if m is not None else ())})
    short = {"text1": "t1", "background1": "bg1", "text2": "t2", "background2": "bg2"}
    drawing = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "hyperlink": "hlink", "followedHyperlink": "folHlink"}

    def expected(name, tint, shade):
        name = mapping.get(short.get(name, name), name)
        rgb = base.get(drawing.get(name, name))
        if not rgb or not re.fullmatch(r"[0-9A-Fa-f]{6}", rgb):
            return None
        if not (tint or shade):
            return rgb.upper()
        h, light, s = colorsys.rgb_to_hls(*(int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4)))
        x = int(tint or shade, 16) / 255
        light = light * x + 1 - x if tint else light * x
        return "".join(f"{round(v * 255):02X}" for v in colorsys.hls_to_rgb(h, light, s))

    n, stale = 0, []
    with zipfile.ZipFile(ctx.output) as z:
        used = _docx_used_styles(z)
        for part in z.namelist():
            if not re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|numbering|styles)\.xml$", part):
                continue
            root = etree.fromstring(z.read(part))
            scopes = [root] if part != "word/styles.xml" else \
                [x for x in root if x.tag == w("docDefaults") or (x.tag == w("style") and x.get(w("styleId")) in used)]
            for el in (e for sc in scopes for e in sc.iter(etree.Element)):
                for ref, ti, sh, cache in (("themeColor", "themeTint", "themeShade", "val" if el.tag == w("color") else "color"),
                                           ("themeFill", "themeFillTint", "themeFillShade", "fill")):
                    name, got = el.get(w(ref)), el.get(w(cache))
                    if not name or name == "none" or not got or got == "auto":
                        continue
                    want = expected(name, el.get(w(ti)), el.get(w(sh)))
                    if want is None:
                        continue
                    n += 1
                    if any(abs(int(got[i:i + 2], 16) - int(want[i:i + 2], 16)) > tol for i in (0, 2, 4)):
                        stale.append(f"{part.split('/')[-1]} {name}: {got.upper()} (theme gives {want})")
    if stale:
        return False, f"{len(stale)}/{n} theme-bound colors in use are stale, e.g. {stale[:3]}"
    return True, f"{n} theme-bound colors in use match the theme"


def check_docx_any_color(ctx, color, scope="headings"):
    """New document: headings (or any text) use the requested color."""
    st = DocxStyles(ctx.output)
    hits = 0
    for para, t, chars in docx_chars(ctx.output):
        if scope == "headings" and st.level(st.para_style(para)) is None:
            continue
        if chars and all(color_ok(docx_effective(st, para, r).get("color"), color) for ch, r in chars if ch.strip()):
            hits += 1
    return hits >= 1, f"{hits} {scope} paragraphs in {color}"


def check_docx_tables(ctx, min_tables=1):
    n = len(_docx_tables(ctx.output))
    return n >= min_tables, f"{n} tables"


# ---- PPTX -----------------------------------------------------------------
def _prs(path):
    import pptx
    return pptx.Presentation(path)


def _slide_by(prs, index=None, title=None):
    slides = list(prs.slides)
    if title:
        for s in slides:
            if s.shapes.title is not None and title.lower() in (s.shapes.title.text or "").lower():
                return s
        return None
    return slides[index - 1] if index and index <= len(slides) else None


def _shape_with(slide, text=None, kind=None):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for sh in slide.shapes:
        for x in ([sh] + (list(sh.shapes) if sh.shape_type == MSO_SHAPE_TYPE.GROUP else [])):
            if kind == "pic" and x.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return x
            if text is not None and x.has_text_frame and text in x.text_frame.text:
                return x
    return None


def _run_props(run_el, theme=None):
    rpr = run_el.find(a("rPr"))
    out = {"bold": False, "italic": False, "underline": False, "strike": False, "size": None, "color": None, "font": None}
    if rpr is None:
        return out
    out["bold"] = rpr.get("b") in ("1", "true")
    out["italic"] = rpr.get("i") in ("1", "true")
    out["underline"] = rpr.get("u") not in (None, "none")
    out["strike"] = rpr.get("strike") not in (None, "noStrike")
    out["size"] = int(rpr.get("sz")) / 100 if rpr.get("sz") else None
    clr = rpr.find(f"{a('solidFill')}/{a('srgbClr')}")
    out["color"] = clr.get("val") if clr is not None else None
    lat = rpr.find(a("latin"))
    out["font"] = lat.get("typeface") if lat is not None else None
    return out


def check_pptx_chars(ctx, text, props, slide=None, title=None, in_title=None):
    prs = _prs(ctx.output)
    sl = _slide_by(prs, slide, title)
    if sl is None:
        return False, "slide not found"
    for sh in sl.shapes:
        if not sh.has_text_frame:
            continue
        is_title = sh.is_placeholder and sh.placeholder_format.type is not None and \
            "TITLE" in str(sh.placeholder_format.type)
        if in_title is not None and is_title != in_title:
            continue
        for para in sh.text_frame.paragraphs:
            ptext = "".join(r.text for r in para.runs)
            i = ptext.find(text)
            if i < 0:
                continue
            pos = 0
            for r in para.runs:
                s, e = pos, pos + len(r.text)
                pos = e
                if e <= i or s >= i + len(text) or not r.text.strip():
                    continue
                got = _run_props(r._r)
                for k, v in props.items():
                    if not _prop_ok(k, got.get(k), v):
                        return False, f"run {r.text!r}: {k}={got.get(k)!r}, expected {v!r}"
            return True, f"{text!r} has {props}"
    return False, f"{text!r} not found on slide"


def check_pptx_text(ctx, slide=None, title=None, has=(), lacks=()):
    prs = _prs(ctx.output)
    sl = _slide_by(prs, slide, title)
    if sl is None:
        return False, "slide not found"
    t = norm(" ".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame) + " " +
             " ".join(c.text for sh in sl.shapes if getattr(sh, "has_table", False) and sh.has_table
                      for r in sh.table.rows for c in r.cells))
    miss = [x for x in has if norm(x) not in t]
    bad = [x for x in lacks if norm(x) in t]
    return not miss and not bad, f"missing {miss} / unexpected {bad}" if miss or bad else "slide text ok"


def _slide_sig(sl):
    out = []
    for sh in sl.shapes:
        el = sh._element
        out.append((sh.shape_id, etree.tostring(el, method="c14n")))
    return out


def check_pptx_others_same(ctx, changed_titles=(), changed_slides=()):
    """Slides other than the changed ones are XML-identical (matched by title)."""
    A_, B_ = _prs(ctx.original), _prs(ctx.output)
    tit = lambda s: (s.shapes.title.text if s.shapes.title is not None else "") or ""
    bmap = {}
    for s in B_.slides:
        bmap.setdefault(tit(s), []).append(s)
    for i, s in enumerate(A_.slides, 1):
        t = tit(s)
        if i in changed_slides or any(c and c.lower() in t.lower() for c in changed_titles):
            continue
        cand = bmap.get(t)
        if not cand:
            return False, f"slide {i} {t!r} missing or retitled"
        if not any(_slide_sig(s) == _slide_sig(x) for x in cand):
            return False, f"slide {i} {t!r} changed"
    return True, "other slides unchanged"


def check_pptx_geom(ctx, slide, text=None, kind=None, pos=None, scale=None, width_factor=None, keep_size=True):
    A_, B_ = _prs(ctx.original), _prs(ctx.output)
    sa, sb = _slide_by(A_, slide), _slide_by(B_, slide)
    x, y = _shape_with(sa, text, kind), _shape_with(sb, text, kind)
    if x is None or y is None:
        return False, "shape not found"
    W, H = B_.slide_width, B_.slide_height
    if pos:
        cx, cy = y.left + y.width / 2, y.top + y.height / 2
        near_r, near_l = W - (y.left + y.width) <= W * 0.15, y.left <= W * 0.15
        near_t, near_b = y.top <= H * 0.15, H - (y.top + y.height) <= H * 0.15
        want = {"top-right": near_r and near_t, "top-left": near_l and near_t,
                "bottom-left": near_l and near_b, "bottom-right": near_r and near_b,
                "center": (abs(cx - W / 2) < W * 0.12 and abs(cy - H / 2) < H * 0.15)}[pos]
        if not want:
            return False, f"shape at ({y.left},{y.top}) size ({y.width},{y.height}) not {pos}"
        if y.left < 0 or y.top < 0 or y.left + y.width > W * 1.001 or y.top + y.height > H * 1.001:
            return False, "shape moved off the slide"
        if keep_size and (abs(x.width - y.width) > 10 or abs(x.height - y.height) > 10):
            return False, "size changed while moving"
    if scale:
        if abs(y.width / x.width - scale) > 0.05 or abs(y.height / x.height - scale) > 0.05:
            return False, f"scale {y.width / x.width:.2f}x{y.height / x.height:.2f}, expected {scale}"
    if width_factor:
        if abs(y.width / x.width - width_factor) > 0.05 or abs(y.height - x.height) > 10:
            return False, f"width x{y.width / x.width:.2f}, height {x.height}->{y.height}"
    return True, f"geometry ok ({y.left},{y.top},{y.width},{y.height})"


def check_pptx_align(ctx, slide, shape_text, align):
    sh = _shape_with(_slide_by(_prs(ctx.output), slide), shape_text)
    if sh is None:
        return False, "shape not found"
    got = []
    for para in sh.text_frame.paragraphs:
        if para.text.strip():
            ppr = para._p.find(a("pPr"))
            got.append(ppr.get("algn") if ppr is not None else None)
    return all(g == align for g in got), f"paragraph alignment {got}"


def check_pptx_fill(ctx, slide, text, color):
    sl = _slide_by(_prs(ctx.output), slide)
    sh = _shape_with(sl, text)
    if sh is None:
        return False, "shape not found"
    clr = sh._element.find(f"{p('spPr')}/{a('solidFill')}/{a('srgbClr')}")
    v = clr.get("val") if clr is not None else None
    return color_ok(v, color), f"fill {v}"


def check_pptx_bg(ctx, title, color, others_keep=True):
    prs = _prs(ctx.output)
    sl = _slide_by(prs, title=title)
    bg = sl._element.find(f"{p('cSld')}/{p('bg')}")
    clr = bg.find(f".//{a('srgbClr')}") if bg is not None else None
    v = clr.get("val") if clr is not None else None
    if not color_ok(v, color):
        return False, f"background {v}"
    with zipfile.ZipFile(ctx.original) as za, zipfile.ZipFile(ctx.output) as zb:
        masters = [n for n in za.namelist() if n.startswith("ppt/slideMasters/") and n.endswith(".xml")]
        if any(za.read(n) != zb.read(n) for n in masters):
            return False, "slide master changed"
    return True, f"background {v}"


def check_pptx_notes(ctx, slide, equals):
    sl = _slide_by(_prs(ctx.output), slide)
    t = sl.notes_slide.notes_text_frame.text if sl is not None and sl.has_notes_slide else ""
    return norm(t) == norm(equals), f"notes {t!r}"


def check_pptx_slides(ctx, count=None, titles=None, hidden=None, visible=None):
    prs = _prs(ctx.output)
    slides = list(prs.slides)
    tit = [((s.shapes.title.text if s.shapes.title is not None else "") or "") for s in slides]
    if count is not None and len(slides) != count:
        return False, f"{len(slides)} slides, expected {count}"
    if titles and tit[:len(titles)] != titles:
        return False, f"titles {tit}"
    for t in hidden or ():
        s = next((s for s, x in zip(slides, tit) if t.lower() in x.lower()), None)
        if s is None or s._element.get("show") not in ("0", "false"):
            return False, f"{t!r} not hidden"
    for t in visible or ():
        s = next((s for s, x in zip(slides, tit) if t.lower() in x.lower()), None)
        if s is None or s._element.get("show") in ("0", "false"):
            return False, f"{t!r} not visible"
    was = ""
    if ctx.original and ctx.original.lower().endswith(".pptx"):
        orig = [((s.shapes.title.text if s.shapes.title is not None else "") or "") for s in _prs(ctx.original).slides]
        was = f" (uploaded file: {orig})" if orig != tit else ""
    return True, f"slides {tit}{was}"


def check_pptx_link(ctx, from_title, to_title):
    prs = _prs(ctx.output)
    src = _slide_by(prs, title=from_title)
    dst = _slide_by(prs, title=to_title)
    if src is None or dst is None:
        return False, "slides not found"
    targets = [r.target_part for r in src.part.rels.values() if r.reltype.endswith("/slide")]
    return dst.part in targets, "link target ok" if dst.part in targets else "link broken"


def _pptx_table(prs, slide):
    sl = _slide_by(prs, slide)
    for sh in sl.shapes:
        if getattr(sh, "has_table", False) and sh.has_table:
            return sh.table
    return None


def check_pptx_table(ctx, slide, row=None, col=None, value=None, cols=None, header=None, row_bold=None,
                     row_align=None, cell_fill=None):
    t = _pptx_table(_prs(ctx.output), slide)
    if t is None:
        return False, "table not found"
    rows = [[c.text for c in r.cells] for r in t.rows]
    seen = f"table ok {rows[0]}"
    if value is not None:
        ci = rows[0].index(col) if col in rows[0] else None
        rr = next((r for r in rows if r[0] == row), None)
        if ci is None or rr is None or rr[ci] != value:
            return False, f"cell {row}/{col} = {None if rr is None or ci is None else rr[ci]!r}"
        seen += f"; row {rr}"
        ot = _pptx_table(_prs(ctx.original), slide) if ctx.original and ctx.original.lower().endswith(".pptx") else None
        orr = next((r for r in ([c.text for c in x.cells] for x in ot.rows) if r[0] == row), None) if ot is not None else None
        if orr and orr != rr:
            seen += f" (uploaded file: {orr})"
    if cols is not None and len(rows[0]) != cols:
        return False, f"{len(rows[0])} columns"
    if header and header not in rows[0]:
        return False, f"header {rows[0]}"
    if row_bold is not None:
        for c in t.rows[row_bold].cells:
            for para in c.text_frame.paragraphs:
                for r in para.runs:
                    if r.text.strip() and not _run_props(r._r)["bold"]:
                        return False, f"{r.text!r} not bold"
    if row_align is not None:
        for c in t.rows[row_align[0]].cells:
            for para in c.text_frame.paragraphs:
                ppr = para._p.find(a("pPr"))
                if ppr is None or ppr.get("algn") != row_align[1]:
                    return False, "header alignment not set"
    if cell_fill:
        rr = next((i for i, r in enumerate(rows) if r[0] == cell_fill[0]), None)
        ci = rows[0].index(cell_fill[1])
        tc = t.rows[rr].cells[ci]._tc
        clr = tc.find(f"{a('tcPr')}/{a('solidFill')}/{a('srgbClr')}")
        if clr is None or not color_ok(clr.get("val"), cell_fill[2]):
            return False, "cell fill missing"
    return True, seen


def check_pptx_theme(ctx, slot=None, color=None, major=None):
    with zipfile.ZipFile(ctx.output) as z:
        themes = [etree.fromstring(z.read(n)) for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
    for th in themes[:1]:
        if slot:
            el = th.find(f".//{a('clrScheme')}/{a(slot)}")
            v = el[0].get("val") if el is not None and len(el) else None
            if not color_ok(v, color):
                return False, f"{slot}={v}"
        if major:
            lat = th.find(f".//{a('majorFont')}/{a('latin')}")
            if lat is None or lat.get("typeface") != major:
                return False, "major font not set"
    return True, "theme ok"


def check_pptx_design_from(ctx, template, min_slides=1, titles_nonempty=True, forbid_titles=()):
    """New deck built on a template: same theme + masters, own slides."""
    with zipfile.ZipFile(template) as za, zipfile.ZipFile(ctx.output) as zb:
        for n in za.namelist():
            if re.match(r"ppt/(theme/theme1|slideMasters/slideMaster1)\.xml$", n):
                if n not in zb.namelist():
                    return False, f"{n} missing"
                ta, tb = etree.fromstring(za.read(n)), etree.fromstring(zb.read(n))
                if n.startswith("ppt/theme") and etree.tostring(ta.find(f".//{a('clrScheme')}"), method="c14n") != \
                        etree.tostring(tb.find(f".//{a('clrScheme')}"), method="c14n"):
                    return False, "template color scheme not kept"
    prs = _prs(ctx.output)
    tit = [((s.shapes.title.text if s.shapes.title is not None else "") or "") for s in prs.slides]
    if len(tit) < min_slides:
        return False, f"{len(tit)} slides"
    if titles_nonempty and any(not x.strip() for x in tit):
        return False, f"empty titles {tit}"
    if any(f.lower() == x.lower() for f in forbid_titles for x in tit):
        return False, f"template example slides left: {tit}"
    return True, f"{len(tit)} slides on template design"


def check_pptx_chart(ctx, min_charts=1):
    with zipfile.ZipFile(ctx.output) as z:
        n = len([x for x in z.namelist() if re.match(r"ppt/charts/chart\d+\.xml$", x)])
    return n >= min_charts, f"{n} charts"


def check_pptx_bullets_max(ctx, maximum=6):
    prs = _prs(ctx.output)
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame and sh != s.shapes.title:
                n = len([pp for pp in sh.text_frame.paragraphs if pp.text.strip()])
                if n > maximum:
                    return False, f"slide {i}: {n} bullets"
    return True, "bullets within limit"


# ---- XLSX -----------------------------------------------------------------
def _wb(path, data_only=False):
    import openpyxl
    return openpyxl.load_workbook(path, data_only=data_only)


def _cell(wb, ref):
    sheet, addr = ref.rsplit("!", 1)
    return wb[sheet.strip("'")][addr]


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return x


def check_xlsx_value(ctx, cell, value):
    v = _cell(_wb(ctx.output), cell).value
    ok = (_num(v) == _num(value)) if isinstance(_num(value), float) else (str(v) == str(value))
    return ok, f"{cell} = {v!r} (expected {value!r})"


def check_xlsx_formula(ctx, cell, contains=None, equals=None, constant=False):
    v = _cell(_wb(ctx.output), cell).value
    if constant:
        return not (isinstance(v, str) and v.startswith("=")), f"{cell} = {v!r}"
    if not (isinstance(v, str) and v.startswith("=")):
        return False, f"{cell} is not a formula: {v!r}"
    if equals and v.replace(" ", "").upper() != equals.replace(" ", "").upper():
        return False, f"{cell} = {v}"
    if contains and not re.search(contains, v, re.I):
        return False, f"{cell} = {v}"
    return True, f"{cell} = {v}"


def check_xlsx_cached(ctx, cell, value, tol=0.01):
    v = _cell(_wb(ctx.output, data_only=True), cell).value
    ok = v is not None and abs(float(v) - float(value)) <= tol if isinstance(_num(v), float) else str(v) == str(value)
    return ok, f"cached {cell} = {v!r} (expected {value})"


_LO_CACHE = {}


def lo_recalc(path):
    if path in _LO_CACHE:
        return _LO_CACHE[path]
    out = Path(tempfile.mkdtemp(prefix="eval-lo-"))
    prof = Path(tempfile.gettempdir()) / "eval-lo-profile"
    (prof / "user").mkdir(parents=True, exist_ok=True)
    xcu = prof / "user" / "registrymodifications.xcu"
    if not xcu.exists():
        xcu.write_text('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                       'xmlns:xs="http://www.w3.org/2001/XMLSchema"><item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
                       '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>')
    subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={prof.as_uri()}", "--convert-to",
                    "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), path], capture_output=True, timeout=180)
    res = out / (Path(path).stem + ".xlsx")
    _LO_CACHE[path] = str(res) if res.exists() else None
    return _LO_CACHE[path]


def check_xlsx_cache_correct(ctx):
    """Every formula's stored result equals LibreOffice's fresh calculation."""
    calc = lo_recalc(ctx.output)
    if not calc:
        return False, "LibreOffice recalculation failed"
    f_wb, got, want = _wb(ctx.output), _wb(ctx.output, True), _wb(calc, True)
    bad, n = [], 0
    for ws in f_wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n += 1
                    g = got[ws.title][c.coordinate].value
                    e = want[ws.title][c.coordinate].value
                    if isinstance(_num(e), float) and isinstance(_num(g), float):
                        if abs(float(g) - float(e)) > 1e-6 * max(1, abs(float(e))):
                            bad.append(f"{ws.title}!{c.coordinate}: {g} vs {e}")
                    elif (g in (None, "") and e not in (None, "")) or (g not in (None, "") and str(g) != str(e)):
                        bad.append(f"{ws.title}!{c.coordinate}: {g!r} vs {e!r}")
    return not bad, f"{n} formulas; stale/missing: {bad[:4]}" if bad else f"{n} formula results correct"


def _rng(ref):
    sheet, addr = ref.rsplit("!", 1)
    from openpyxl.utils.cell import range_boundaries
    c1, r1, c2, r2 = range_boundaries(addr)
    return sheet.strip("'"), c1, r1, c2, r2


def _cells(wb, ref):
    sheet, c1, r1, c2, r2 = _rng(ref)
    ws = wb[sheet]
    return [ws.cell(row=r, column=c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def check_xlsx_style(ctx, range, font=None, fill=None, align=None, numfmt=None, keep=None):
    wb = _wb(ctx.output)
    for c in _cells(wb, range):
        f = c.font
        for k, v in (font or {}).items():
            got = {"bold": f.b, "italic": f.i, "underline": f.u not in (None, "none"), "strike": f.strike,
                   "size": f.sz, "font": f.name, "color": (f.color.rgb if f.color is not None and f.color.type == "rgb" else None)}[k]
            if not _prop_ok(k, got, v):
                return False, f"{c.coordinate} font {k}={got!r}, expected {v!r}"
        if fill:
            fg = c.fill.fgColor.rgb if c.fill is not None and c.fill.fill_type == "solid" and c.fill.fgColor.type == "rgb" else None
            if not color_ok(fg, fill):
                return False, f"{c.coordinate} fill {fg}"
        if align and (c.alignment.horizontal or "general") != align:
            return False, f"{c.coordinate} align {c.alignment.horizontal}"
        if numfmt and not re.search(numfmt, c.number_format or ""):
            return False, f"{c.coordinate} number format {c.number_format!r}"
    if keep:
        o = _wb(ctx.original)
        for c in _cells(wb, range):
            oc = o[c.parent.title][c.coordinate]
            for k in keep:
                got = {"bold": (c.font.b, oc.font.b), "italic": (c.font.i, oc.font.i),
                       "underline": (c.font.u, oc.font.u), "fill": (c.fill.fgColor.rgb, oc.fill.fgColor.rgb),
                       "border": (repr(c.border), repr(oc.border)), "color": (c.font.color and c.font.color.rgb,
                                                                              oc.font.color and oc.font.color.rgb)}[k]
                if bool(got[0]) != bool(got[1]) and k != "fill" or (k in ("fill", "color", "border") and got[0] != got[1]):
                    return False, f"{c.coordinate}: {k} changed {got[1]!r} -> {got[0]!r}"
    return True, f"{range} style ok"


def check_xlsx_cf(ctx, sheet, sqref_has, min_rules=1, keep_from_original=True, exclude=()):
    """A rule covers ``sqref_has`` (one cell, or a list of which any one will
    do); no rule added by the agent covers a cell in ``exclude``."""
    wb = _wb(ctx.output)
    ws = wb[sheet]
    rules = [(str(rng.sqref), r) for rng in ws.conditional_formatting for r in rng.rules]
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, range_boundaries

    def covers(sq, cell):
        col, row = coordinate_from_string(cell)
        ci = column_index_from_string(col)
        for part in sq.split():
            c1, r1, c2, r2 = range_boundaries(part if ":" in part else f"{part}:{part}")
            if c1 <= ci <= c2 and r1 <= row <= r2:
                return True
        return False
    seen = [f"{sq} {r.type} {r.operator or ''} {' '.join(r.formula or [])}".replace("  ", " ").strip() for sq, r in rules]
    want = [sqref_has] if isinstance(sqref_has, str) else sqref_has
    hit = [x for x in rules if any(covers(x[0], c) for c in want)]
    if len(hit) < min_rules:
        return False, f"rules {seen}"
    o = [str(r.sqref) for r in _wb(ctx.original)[sheet].conditional_formatting] if ctx.original else []
    if keep_from_original and any(x not in [y[0] for y in rules] for x in o):
        return False, "original conditional format lost"
    over = sorted({c for sq, _ in rules if sq not in o for c in exclude if covers(sq, c)})
    if over:
        return False, f"rules {seen} also cover {over}"
    return True, f"rules {seen}"


def check_xlsx_width(ctx, sheet, col, minimum):
    wb = _wb(ctx.output)
    wdt = wb[sheet].column_dimensions[col].width
    return (wdt or 0) >= minimum, f"{sheet}!{col} width {wdt}"


def xlsx_features(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        out["charts"] = len([n for n in names if re.match(r"xl/charts/chart\d+\.xml$", n)])
        out["images"] = len([n for n in names if n.startswith("xl/media/")])
        out["external_links"] = len([n for n in names if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n)])
        out["tables"] = sorted(etree.fromstring(z.read(n)).get("name") for n in names if re.match(r"xl/tables/table\d+\.xml$", n))
        wbx = etree.fromstring(z.read("xl/workbook.xml"))
        out["names"] = sorted(d.get("name") for d in wbx.iter(f"{{{S}}}definedName"))
        out["sheets"] = [s.get("name") for s in wbx.iter(f"{{{S}}}sheet")]
        cf = dv = mc = 0
        for n in names:
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                ws = etree.fromstring(z.read(n))
                cf += len(ws.findall(f"{{{S}}}conditionalFormatting"))
                dv += len(list(ws.iter(f"{{{S}}}dataValidation")))
                mc += len(list(ws.iter(f"{{{S}}}mergeCell")))
        out.update(cf=cf, dv=dv, merged=mc)
    return out


def check_xlsx_features_kept(ctx, cf_delta=0):
    a_, b_ = xlsx_features(ctx.original), xlsx_features(ctx.output)
    bad = []
    for k, v in a_.items():
        if k == "cf":
            if b_[k] < v + cf_delta:
                bad.append(f"cf {v}->{b_[k]}")
        elif isinstance(v, list):
            if not set(v) <= set(b_[k]):
                bad.append(f"{k} {v}->{b_[k]}")
        elif b_[k] < v:
            bad.append(f"{k} {v}->{b_[k]}")
    return not bad, f"lost {bad}" if bad else "tables/CF/validation/names/charts/images/links kept"


def check_xlsx_table_ref(ctx, name, ref):
    with zipfile.ZipFile(ctx.output) as z:
        for n in z.namelist():
            if re.match(r"xl/tables/table\d+\.xml$", n):
                t = etree.fromstring(z.read(n))
                if t.get("name") == name:
                    return t.get("ref") == ref, f"table {name} ref {t.get('ref')}"
    return False, f"table {name} missing"


def check_xlsx_others_same(ctx, except_=(), shifted=None):
    """Values, formulas and visible formats of cells outside ``except_``
    ranges are unchanged (cached formula results are not compared here).
    shifted = {"sheet": S, "at": row, "n": count}: compare shifted rows."""
    from openpyxl.utils.cell import range_boundaries
    a_, b_ = _wb(ctx.original), _wb(ctx.output)
    ex = []
    for r in except_:
        sh, addr = r.rsplit("!", 1)
        ex.append((sh.strip("'"), range_boundaries(addr) if ":" in addr else range_boundaries(f"{addr}:{addr}")))
    inside = lambda sh, c: any(sh == s and b[0] <= c.column <= b[2] and b[1] <= c.row <= b[3] for s, b in ex)
    sig = lambda c: (c.value, c.font.b, c.font.i, c.font.u, c.fill.fgColor.rgb if c.fill.fill_type else None,
                     c.number_format, c.alignment.horizontal)
    for ws in a_:
        if ws.title not in b_.sheetnames:
            return False, f"sheet {ws.title} missing"
        wb_ = b_[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if inside(ws.title, c):
                    continue
                r = c.row
                if shifted and ws.title == shifted["sheet"] and r >= shifted["at"]:
                    r += shifted["n"]
                if r < 1:
                    continue
                d = wb_.cell(row=r, column=c.column)
                if isinstance(c.value, str) and c.value.startswith("="):
                    continue  # formulas: checked by dedicated checks (references may shift)
                if sig(c) != sig(d):
                    return False, f"{ws.title}!{c.coordinate} changed {sig(c)[:2]} -> {sig(d)[:2]}"
    return True, "other cells unchanged"


def check_xlsx_sorted(ctx, sheet, col, first, last, desc=False, keep_rows_from=None):
    wb = _wb(ctx.output, data_only=True)
    ws = wb[sheet]
    vals = [ws[f"{col}{r}"].value for r in range(first, last + 1)]
    nums = [v for v in vals if v is not None]
    ok = nums == sorted(nums, reverse=desc, key=lambda v: (str(type(v)), v if not isinstance(v, str) else v.casefold()))
    if not ok:
        return False, f"not sorted: {nums[:6]}"
    if keep_rows_from:
        cols = keep_rows_from
        o = _wb(ctx.original)[sheet]
        n = _wb(ctx.output)[sheet]
        rows = lambda w_: sorted(tuple(str(w_[f"{c}{r}"].value) for c in cols) for r in range(first, last + 1))
        if rows(o) != rows(n):
            return False, "row contents not kept together"
    return True, "sorted, rows kept together"


def check_xlsx_dv(ctx, sheet, min_rules=1):
    ws = _wb(ctx.output)[sheet] if sheet else _wb(ctx.output).active
    n = len(ws.data_validations.dataValidation)
    return n >= min_rules, f"{n} data validations"


def check_xlsx_any_formula(ctx, pattern, min_count=1):
    wb = _wb(ctx.output)
    n = sum(1 for ws in wb for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("=") and re.search(pattern, c.value, re.I))
    return n >= min_count, f"{n} formulas matching {pattern}"


# ---- PDF ------------------------------------------------------------------
def pdf_page_text(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        return [pg.extract_text() or "" for pg in pdf.pages]


def pdf_lines(path):
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            ls = []
            for d in pg.extract_text_lines(return_chars=True, strip=True):
                ch = next((c for c in d["chars"] if c["text"].strip()), None)
                ls.append({"text": d["text"], "top": d["top"], "x0": d["x0"],
                           "font": re.sub(r"^[A-Z]{6}\+", "", ch["fontname"]) if ch else None,
                           "size": round(ch["size"], 1) if ch else None, "chars": d["chars"]})
            out.append(ls)
    return out


def check_pdf_pages(ctx, n):
    k = len(pdf_page_text(ctx.output))
    was = f", uploaded file had {len(pdf_page_text(ctx.original))}" if ctx.original and ctx.original.lower().endswith(".pdf") else ""
    return k == n, f"{k} pages (expected {n}{was})"


def check_pdf_page_has(ctx, page, has=(), lacks=()):
    t = pdf_page_text(ctx.output)
    if page > len(t):
        return False, "page missing"
    pt = norm(t[page - 1])
    miss = [x for x in has if norm(x) not in pt]
    bad = [x for x in lacks if norm(x) in pt]
    return not miss and not bad, (f"page {page}: missing {miss} unexpected {bad}" if miss or bad else
                                  f"page {page} has {list(has)}" + (f", lacks {list(lacks)}" if lacks else ""))


def pdf_body_text(path, band=0.08):
    """Page texts without the header/footer band (page numbers may change)."""
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            h = pg.height
            out.append("\n".join(d["text"] for d in pg.extract_text_lines(strip=True)
                                 if d["bottom"] >= h * band and d["top"] <= h * (1 - band)))
    return out


def check_pdf_pages_same(ctx, mapping):
    """mapping {result_page: original_page}: identical body text."""
    a_, b_ = pdf_body_text(ctx.original), pdf_body_text(ctx.output)
    for rp, op in mapping.items():
        rp, op = int(rp), int(op)
        if rp > len(b_) or norm(b_[rp - 1]) != norm(a_[op - 1]):
            return False, f"result page {rp} != original page {op}"
    return True, "page order/content as expected: " + ", ".join(f"new p{k} = original p{v}" for k, v in mapping.items())


def check_pdf_page_others_same(ctx, page, changed=()):
    """On ``page`` every line except those containing ``changed`` texts keeps
    its text and position (±1pt)."""
    A_, B_ = pdf_lines(ctx.original)[page - 1], pdf_lines(ctx.output)[page - 1]
    keep = [l for l in A_ if not any(c in l["text"] for c in changed)]
    for l in keep:
        m = next((x for x in B_ if x["text"] == l["text"]), None)
        if m is None:
            return False, f"line {l['text'][:50]!r} changed or missing"
        if abs(m["top"] - l["top"]) > 1 or abs(m["x0"] - l["x0"]) > 1:
            return False, f"line {l['text'][:40]!r} moved"
    return True, "other lines unchanged"


def check_pdf_font_kept(ctx, page, new_text, like_text):
    """The new text is drawn with the same font family and size as the
    original line it replaced."""
    A_, B_ = pdf_lines(ctx.original)[page - 1], pdf_lines(ctx.output)[page - 1]
    o = next((l for l in A_ if like_text in l["text"]), None)
    n = next((l for l in B_ if norm(new_text)[:20] in norm(l["text"])), None)
    if o is None or n is None:
        return False, "lines not found"
    # metric-compatible families count as the same look
    same = {"helvetica": "liberationsans", "arial": "liberationsans", "arialmt": "liberationsans",
            "times": "liberationserif", "timesnewroman": "liberationserif", "calibri": "carlito", "cambria": "caladea"}
    fam = lambda f: (lambda x: same.get(x, x))(re.split(r"[-,]", f or "")[0].replace(" ", "").lower())
    ok = fam(o["font"]) == fam(n["font"]) and abs((o["size"] or 0) - (n["size"] or 0)) < 0.3
    return ok, f"font {o['font']} {o['size']} -> {n['font']} {n['size']}"


def check_pdf_row_has(ctx, page, key, value):
    """The table row (text line) starting with ``key`` contains ``value``."""
    for line in pdf_page_text(ctx.output)[page - 1].splitlines():
        if line.startswith(key) or f" {key} " in f" {line} ":
            ok = value in line.split() or value in line
            return ok, f"row {line!r}"
    return False, f"row {key!r} not found"


def check_pdf_footer(ctx, pattern):
    t = pdf_page_text(ctx.output)
    n = len(t)
    for i, pt in enumerate(t, 1):
        want = pattern.format(n=i, total=n)
        if want not in pt:
            return False, f"page {i} footer lacks {want!r}"
    return True, f"page numbers consistent: {pattern.format(n=1, total=n)!r} … {pattern.format(n=n, total=n)!r}"


def check_pdf_color(ctx, text, color, page=1):
    import pdfplumber
    with pdfplumber.open(ctx.output) as pdf:
        pg = pdf.pages[page - 1]
        line = next((d for d in pg.extract_text_lines(return_chars=True) if text.lower() in d["text"].lower()), None)
        if line is None:
            return False, f"{text!r} not found"
        c = next(ch for ch in line["chars"] if ch["text"].strip())
        col = c.get("non_stroking_color")
        if isinstance(col, (list, tuple)) and len(col) == 3:
            hexv = "".join(f"{round(x * 255):02X}" for x in col)
        elif isinstance(col, (list, tuple)) and len(col) == 4:
            k = col[3]
            hexv = "".join(f"{round(255 * (1 - x) * (1 - k)):02X}" for x in col[:3])
        else:
            hexv = "000000"
        return color_ok(hexv, color), f"{text!r} color {hexv}"


def check_pdf_tables(ctx, min_tables=1):
    import pdfplumber
    with pdfplumber.open(ctx.output) as pdf:
        n = sum(len(pg.find_tables()) for pg in pdf.pages)
    return n >= min_tables, f"{n} tables"


def check_pdf_landscape(ctx):
    import pdfplumber
    with pdfplumber.open(ctx.output) as pdf:
        pg = pdf.pages[0]
        return pg.width > pg.height, f"{pg.width:.0f}x{pg.height:.0f}"


def check_pdf_rotated(ctx, page, degrees):
    import pypdfium2 as pdfium
    rot = pdfium.PdfDocument(ctx.output)[page - 1].get_rotation()
    return rot % 360 == degrees % 360, f"page {page} rotation {rot}"


def check_pdf_every_page_has(ctx, text):
    """Text on every page (pdfium extraction, whitespace-insensitive, so
    rotated watermark text counts too)."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(ctx.output)
    flat = lambda s: re.sub(r"\s", "", s or "")
    plumber = pdf_page_text(ctx.output)
    miss = []
    for i in range(len(doc)):
        got = flat(doc[i].get_textpage().get_text_range())
        if flat(text) not in got and flat(text) not in flat(plumber[i]):
            miss.append(i + 1)
    return not miss, f"pages without {text!r}: {miss}" if miss else f"{text!r} on every page"


CHECKS = {k[6:]: v for k, v in globals().items() if k.startswith("check_")}
