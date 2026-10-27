"""In-place DOCX editing on top of docgen.package.Package.

    from docgen.package import Package
    from docgen import docx, one
    pkg = Package(src)
    m = one(docx.find(pkg, "risk", section="Ek Notlar", whole_word=True), "risk")
    docx.set_run_props(docx.split_runs(m), bold=True)
    pkg.save(dst)

Every paragraph carries its location: ``part`` (document/header/footer),
``table`` = (table index, row, col) or None, and ``path`` = the chain of
heading texts above it (used for ``section=`` filters).
"""
import copy
import re

from lxml import etree

from . import NotFound, NotSupported
from .package import MC_NS

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MAIN = "word/document.xml"


def w(tag):
    return f"{{{W_NS}}}{tag}"


RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
             "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
             "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath", "rPrChange"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
             "rPr", "sectPr", "pPrChange"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel",
              "cellMerge", "tcPrChange"]
ALIGN = {"left": "left", "center": "center", "centre": "center", "right": "right",
         "justify": "both", "both": "both"}


def put_child(parent, tag, order, attrs=None):
    """Create/replace child ``tag`` at its schema position; returns it."""
    old = parent.find(w(tag))
    if old is not None:
        parent.remove(old)
    el = etree.Element(w(tag))
    for k, v in (attrs or {}).items():
        el.set(w(k), str(v))
    idx = order.index(tag)
    for i, ch in enumerate(parent):
        name = etree.QName(ch).localname
        if name in order and order.index(name) > idx:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def drop_child(parent, tag):
    for el in parent.findall(w(tag)):
        parent.remove(el)


def ensure_first(parent, tag):
    el = parent.find(w(tag))
    if el is None:
        el = etree.Element(w(tag))
        parent.insert(0, el)
    return el


def _in_fallback(el):
    p = el.getparent()
    while p is not None:
        if p.tag == f"{{{MC_NS}}}Fallback" or p.tag == w("del"):
            return True
        p = p.getparent()
    return False


# ---- styles -------------------------------------------------------------
class Styles:
    def __init__(self, pkg):
        self.root = pkg.xml("word/styles.xml") if pkg.has("word/styles.xml") else None
        self.by_id = {}
        if self.root is not None:
            for s in self.root.iter(w("style")):
                self.by_id[s.get(w("styleId"))] = s
        self.default_para = next((sid for sid, s in self.by_id.items()
                                  if s.get(w("type")) == "paragraph" and s.get(w("default")) in ("1", "true")), None)

    def name(self, sid):
        s = self.by_id.get(sid)
        n = s.find(w("name")) if s is not None else None
        return n.get(w("val")) if n is not None else sid

    def chain(self, sid):
        seen = []
        while sid and sid in self.by_id and sid not in seen:
            seen.append(sid)
            b = self.by_id[sid].find(w("basedOn"))
            sid = b.get(w("val")) if b is not None else None
        return seen

    def level(self, sid):
        """Heading level (0 = Title, 1..9) or None."""
        for s in self.chain(sid):
            name = (self.name(s) or "").lower()
            if name == "title":
                return 0
            m = re.fullmatch(r"heading (\d)", name)
            if m:
                return int(m.group(1))
            ppr = self.by_id[s].find(w("pPr"))
            ol = ppr.find(w("outlineLvl")) if ppr is not None else None
            if ol is not None and ol.get(w("val")) not in (None, "9"):
                return int(ol.get(w("val"))) + 1
        return None

    def run_flag(self, sids, tag):
        """True if any style in the chains sets an on/off run property."""
        for sid in sids:
            for s in self.chain(sid):
                rpr = self.by_id[s].find(w("rPr"))
                el = rpr.find(w(tag)) if rpr is not None else None
                if el is not None:
                    return el.get(w("val"), "true") not in ("0", "false", "none")
        return False


# ---- paragraph model ----------------------------------------------------
class Run:
    __slots__ = ("el", "text", "start", "end")

    def __init__(self, el, text, start):
        self.el, self.text, self.start, self.end = el, text, start, start + len(text)


class Para:
    def __init__(self, el, part, index, table, styles):
        self.el, self.part, self.index, self.table = el, part, index, table
        ppr = el.find(w("pPr"))
        ps = ppr.find(w("pStyle")) if ppr is not None else None
        self.style = ps.get(w("val")) if ps is not None else styles.default_para
        self.level = styles.level(self.style) if self.style else None
        ol = ppr.find(w("outlineLvl")) if ppr is not None else None
        if self.level is None and ol is not None and ol.get(w("val")) != "9":
            self.level = int(ol.get(w("val"))) + 1
        self.path = ()
        self._load_runs()

    def _load_runs(self):
        self.runs = []
        pos = 0
        for r in self.el.iter(w("r")):
            if _in_fallback(r):
                continue
            t = "".join(x.text or "" for x in r if x.tag == w("t"))
            self.runs.append(Run(r, t, pos))
            pos += len(t)
        self.text = "".join(r.text for r in self.runs)

    @property
    def is_heading(self):
        return self.level is not None

    def where(self):
        loc = self.part.split("/")[-1]
        if self.table:
            loc += f" table {self.table[0] + 1} row {self.table[1] + 1} col {self.table[2] + 1}"
        sec = " > ".join(self.path) if self.path else "(no heading)"
        return f"{loc} | {sec} | para {self.index}"

    def __repr__(self):
        return f"<Para {self.where()} {self.text[:60]!r}>"


def _table_pos(p, body):
    """(body-level table index, row, col) of the innermost cell, or None."""
    tc = next((a for a in p.iterancestors(w("tc"))), None)
    if tc is None:
        return None
    tr = tc.getparent()
    outer = [a for a in p.iterancestors(w("tbl")) if a.getparent() is body]
    tables = [t for t in body if t.tag == w("tbl")]
    ti = tables.index(outer[0]) if outer else -1
    return (ti, tr.getparent().findall(w("tr")).index(tr), tr.findall(w("tc")).index(tc))


def paragraphs(pkg, include_headers=True):
    styles = Styles(pkg)
    out = []
    body = pkg.xml(MAIN).find(w("body"))
    stack = []
    for i, p in enumerate(x for x in body.iter(w("p")) if not _in_fallback(x)):
        para = Para(p, MAIN, i, _table_pos(p, body), styles)
        if para.level is not None and para.text.strip():
            stack = [s for s in stack if s[0] < para.level] + [(para.level, para.text.strip())]
        para.path = tuple(s[1] for s in stack)
        out.append(para)
    if include_headers:
        for r in pkg.rels(MAIN):
            if r["Type"].endswith(("/header", "/footer")) and pkg.has(r["part"]):
                for i, p in enumerate(x for x in pkg.xml(r["part"]).iter(w("p")) if not _in_fallback(x)):
                    out.append(Para(p, r["part"], i, None, styles))
    return out


def headings(pkg):
    return [p for p in paragraphs(pkg, include_headers=False) if p.is_heading and p.text.strip()]


def _section_ok(para, section):
    if not section:
        return True
    want = [s.strip().casefold() for s in re.split(r"\s*[>/]\s*", section) if s.strip()]
    have = [h.casefold() for h in para.path]
    i = 0
    for h in have:
        if i < len(want) and want[i] in h:
            i += 1
    return i == len(want)


class Match:
    def __init__(self, para, start, end):
        self.para, self.start, self.end = para, start, end

    @property
    def text(self):
        return self.para.text[self.start:self.end]

    def __repr__(self):
        t = self.para.text
        ctx = t[max(0, self.start - 25):self.start] + "[" + self.text + "]" + t[self.end:self.end + 25]
        return f"{self.para.where()} | …{ctx}…"


def find(pkg, text, *, section=None, part=None, in_table=None, whole_word=False,
         ignore_case=False, headings_only=False, include_headers=True):
    """All occurrences of ``text`` with their location. Filters:
    section='Bulgular > Sonuç' (heading path), part='header'/'footer'/'document',
    in_table=True/False, headings_only=True (match heading paragraphs only)."""
    flags = re.IGNORECASE if ignore_case else 0
    pat = re.escape(text)
    if whole_word:
        pat = rf"(?<!\w){pat}(?!\w)"
    rx = re.compile(pat, flags)
    out = []
    for p in paragraphs(pkg, include_headers=include_headers):
        if part and part not in p.part:
            continue
        if in_table is not None and bool(p.table) != in_table:
            continue
        if headings_only and not p.is_heading:
            continue
        if not _section_ok(p, section):
            continue
        out.extend(Match(p, m.start(), m.end()) for m in rx.finditer(p.text))
    return out


def find_paragraphs(pkg, contains=None, *, section=None, style=None, headings_only=False):
    out = []
    for p in paragraphs(pkg):
        if contains is not None and contains not in p.text:
            continue
        if style and p.style != style:
            continue
        if headings_only and not p.is_heading:
            continue
        if _section_ok(p, section):
            out.append(p)
    return out


# ---- run splitting ------------------------------------------------------
def _split_run(r, k):
    """Split run element at text offset k; returns (left, right)."""
    right = copy.deepcopy(r)
    pos = 0
    for lc, rc in zip(list(r), list(right)):
        if lc.tag == w("rPr"):
            continue
        if lc.tag == w("t"):
            t = lc.text or ""
            a, b = pos, pos + len(t)
            if b <= k:
                right.remove(rc)
            elif a >= k:
                r.remove(lc)
            else:
                lc.text, rc.text = t[:k - a], t[k - a:]
                for el in (lc, rc):
                    el.set(XML_SPACE, "preserve")
            pos = b
        else:
            (right.remove(rc) if pos < k else r.remove(lc))
    r.addnext(right)
    return r, right


def split_runs(match_or_para, start=None, end=None):
    """Return the run elements that cover exactly [start, end) of the
    paragraph text, splitting runs at the boundaries (formatting copied)."""
    if isinstance(match_or_para, Match):
        para, start, end = match_or_para.para, match_or_para.start, match_or_para.end
    else:
        para = match_or_para
        start = 0 if start is None else start
        end = len(para.text) if end is None else end
    out = []
    for run in para.runs:
        if run.end <= start or run.start >= end or not run.text:
            continue
        el = run.el
        if run.start < start:
            _, el = _split_run(el, start - run.start)
        if run.end > end:
            el, _ = _split_run(el, end - max(run.start, start))
        out.append(el)
    para._load_runs()
    return out


# ---- formatting ---------------------------------------------------------
def _onoff(rpr, tag, value, style_on):
    cs = tag + "Cs" if tag in ("b", "i") else None
    if value:
        put_child(rpr, tag, RPR_ORDER)
        if cs:
            put_child(rpr, cs, RPR_ORDER)
    else:
        drop_child(rpr, tag)
        if cs:
            drop_child(rpr, cs)
        if style_on:  # inherited from a style: switch off explicitly
            put_child(rpr, tag, RPR_ORDER, {"val": "0"})


def set_run_props(runs, *, pkg=None, bold=None, italic=None, underline=None, strike=None,
                  color=None, size=None, font=None, highlight=None):
    """Set character formatting on run elements. ``None`` = leave as is,
    ``False`` = remove (switched off explicitly if a style turns it on).
    color='C00000', size=11 (pt), font='Arial', underline=True|'double',
    highlight='yellow'. Pass ``pkg`` so inherited style formatting is known."""
    styles = Styles(pkg) if pkg is not None else None
    for r in runs:
        rpr = ensure_first(r, "rPr")
        sids = []
        if styles is not None:
            rs = rpr.find(w("rStyle"))
            if rs is not None:
                sids.append(rs.get(w("val")))
            p = r.getparent()
            while p is not None and p.tag != w("p"):
                p = p.getparent()
            ppr = p.find(w("pPr")) if p is not None else None
            ps = ppr.find(w("pStyle")) if ppr is not None else None
            sids.append(ps.get(w("val")) if ps is not None else styles.default_para)
        on = (lambda tag: styles.run_flag([s for s in sids if s], tag)) if styles else (lambda tag: False)
        if bold is not None:
            _onoff(rpr, "b", bold, on("b"))
        if italic is not None:
            _onoff(rpr, "i", italic, on("i"))
        if strike is not None:
            _onoff(rpr, "strike", strike, on("strike"))
        if underline is not None:
            if underline:
                put_child(rpr, "u", RPR_ORDER, {"val": "single" if underline is True else underline})
            else:
                drop_child(rpr, "u")
                if on("u"):
                    put_child(rpr, "u", RPR_ORDER, {"val": "none"})
        if color is not None:
            if color:
                put_child(rpr, "color", RPR_ORDER, {"val": color.lstrip("#").upper()})
            else:
                drop_child(rpr, "color")
        if size is not None:
            if size:
                hp = str(int(round(float(size) * 2)))
                put_child(rpr, "sz", RPR_ORDER, {"val": hp})
                put_child(rpr, "szCs", RPR_ORDER, {"val": hp})
            else:
                drop_child(rpr, "sz")
                drop_child(rpr, "szCs")
        if font is not None:
            if font:
                put_child(rpr, "rFonts", RPR_ORDER, {"ascii": font, "hAnsi": font, "cs": font, "eastAsia": font})
            else:
                drop_child(rpr, "rFonts")
        if highlight is not None:
            if highlight:
                put_child(rpr, "highlight", RPR_ORDER, {"val": highlight})
            else:
                drop_child(rpr, "highlight")
        if len(rpr) == 0:
            r.remove(rpr)


def _twips(pt):
    return str(int(round(float(pt) * 20)))


def set_paragraph_props(paras, *, align=None, space_before=None, space_after=None,
                        line_spacing=None, indent_left_cm=None, indent_first_cm=None, style=None):
    """Direct paragraph formatting on the given paragraphs only.
    align: left|center|right|justify; space_before/after in pt;
    line_spacing as a multiple (1.5); indents in cm; style = styleId."""
    for p in paras:
        el = p.el if isinstance(p, Para) else p
        ppr = ensure_first(el, "pPr")
        if style is not None:
            put_child(ppr, "pStyle", PPR_ORDER, {"val": style}) if style else drop_child(ppr, "pStyle")
        if align is not None:
            put_child(ppr, "jc", PPR_ORDER, {"val": ALIGN[align]})
        if any(v is not None for v in (space_before, space_after, line_spacing)):
            sp = ppr.find(w("spacing"))
            if sp is None:
                sp = put_child(ppr, "spacing", PPR_ORDER)
            if space_before is not None:
                sp.set(w("before"), _twips(space_before))
            if space_after is not None:
                sp.set(w("after"), _twips(space_after))
            if line_spacing is not None:
                sp.set(w("line"), str(int(round(float(line_spacing) * 240))))
                sp.set(w("lineRule"), "auto")
        if indent_left_cm is not None or indent_first_cm is not None:
            ind = ppr.find(w("ind"))
            if ind is None:
                ind = put_child(ppr, "ind", PPR_ORDER)
            if indent_left_cm is not None:
                ind.set(w("left"), str(int(round(indent_left_cm * 567))))
            if indent_first_cm is not None:
                ind.set(w("firstLine"), str(int(round(indent_first_cm * 567))))


STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
               "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose",
               "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]


def set_style_props(pkg, style_id, *, align=None, space_before=None, space_after=None,
                    line_spacing=None, **run_props):
    """Change a named style in styles.xml — affects every paragraph that uses
    it (scope "all headings of level N" / "all body text")."""
    styles = Styles(pkg)
    s = styles.by_id.get(style_id)
    if s is None:
        raise NotFound(f"style {style_id!r} not found; available: {sorted(styles.by_id)}")
    if any(v is not None for v in (align, space_before, space_after, line_spacing)):
        holder = etree.Element(w("p"))
        old = s.find(w("pPr"))
        if old is not None:
            s.remove(old)
            holder.append(old)
        set_paragraph_props([holder], align=align, space_before=space_before,
                            space_after=space_after, line_spacing=line_spacing)
        new = holder.find(w("pPr"))
        put_child(s, "pPr", STYLE_ORDER).getparent().replace(s.find(w("pPr")), new)
    if run_props:
        holder = etree.Element(w("r"))
        old = s.find(w("rPr"))
        if old is not None:
            s.remove(old)
            holder.append(old)
        set_run_props([holder], **run_props)
        new = holder.find(w("rPr"))
        if new is not None:
            put_child(s, "rPr", STYLE_ORDER).getparent().replace(s.find(w("rPr")), new)


def style_ids_for_level(pkg, level):
    """Style ids whose heading level is ``level`` (incl. derived styles)."""
    st = Styles(pkg)
    return [sid for sid, el in st.by_id.items()
            if el.get(w("type")) == "paragraph" and st.level(sid) == level]


# ---- text edits ---------------------------------------------------------
def _rpr_sig(r):
    rpr = r.find(w("rPr"))
    return etree.tostring(rpr, method="c14n") if rpr is not None else b""


def replace(match, new_text):
    """Replace exactly one match. The new text takes the formatting of the
    first run of the match; returns a warning if the match spanned runs with
    different formatting (report it to the user)."""
    runs = split_runs(match)
    if not runs:
        raise NotFound("match has no text runs")
    warn = None
    if len({_rpr_sig(r) for r in runs}) > 1:
        warn = (f"'{match.text}' spanned differently formatted runs; the replacement "
                f"uses the formatting of '{runs[0].findtext(w('t'))}'")
    first = runs[0]
    ts = first.findall(w("t"))
    ts[0].text = new_text
    ts[0].set(XML_SPACE, "preserve")
    for t in ts[1:]:
        first.remove(t)
    for r in runs[1:]:
        r.getparent().remove(r)
    match.para._load_runs()
    return warn


def replace_all(pkg, old, new, **find_kwargs):
    """Replace every match of ``find(pkg, old, **find_kwargs)``. Use
    whole_word=True to avoid hitting inflected words ('risk' in 'riski').
    Returns (count, warnings)."""
    matches = find(pkg, old, **find_kwargs)
    warns = []
    # within a paragraph go right-to-left so earlier offsets stay valid
    for m in sorted(matches, key=lambda m: (m.para.part, m.para.index, -m.start)):
        wmsg = replace(m, new)
        if wmsg:
            warns.append(wmsg)
    return len(matches), warns


def set_paragraph_text(para, text):
    """Replace the whole text of a paragraph, keeping pPr and the first run's
    formatting (drawings/fields in the paragraph are refused)."""
    el = para.el if isinstance(para, Para) else para
    if el.find(".//" + w("drawing")) is not None or el.find(".//" + w("fldChar")) is not None:
        raise NotSupported("paragraph contains a drawing or field; edit runs individually")
    runs = [r for r in el.iter(w("r")) if not _in_fallback(r)]
    keep = copy.deepcopy(runs[0].find(w("rPr"))) if runs and runs[0].find(w("rPr")) is not None else None
    for ch in list(el):
        if ch.tag != w("pPr"):
            el.remove(ch)
    r = etree.SubElement(el, w("r"))
    if keep is not None:
        r.append(keep)
    t = etree.SubElement(r, w("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return el


def new_paragraph(text, like=None, style=None):
    """Build a paragraph element. ``like`` copies pPr and the first run's rPr
    of an existing paragraph (so it matches the document's body style)."""
    p = etree.Element(w("p"))
    src = like.el if isinstance(like, Para) else like
    if src is not None and src.find(w("pPr")) is not None:
        ppr = copy.deepcopy(src.find(w("pPr")))
        for t in ("sectPr", "numPr", "rPr"):
            drop_child(ppr, t)
        p.append(ppr)
    if style:
        put_child(ensure_first(p, "pPr"), "pStyle", PPR_ORDER, {"val": style})
    r = etree.SubElement(p, w("r"))
    if src is not None:
        fr = next((x for x in src.iter(w("r")) if x.find(w("rPr")) is not None), None)
        if fr is not None:
            r.append(copy.deepcopy(fr.find(w("rPr"))))
    t = etree.SubElement(r, w("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return p


def _block(el):
    """Body-level block (p or tbl) containing ``el``."""
    while el.getparent() is not None and el.getparent().tag != w("body"):
        el = el.getparent()
    return el


def section_blocks(pkg, heading):
    """Body-level elements of a section: the heading paragraph and everything
    after it up to the next heading of the same or a higher level."""
    if not isinstance(heading, Para) or not heading.is_heading:
        raise NotFound("section_blocks needs a heading paragraph (see headings())")
    styles = Styles(pkg)
    out = [heading.el]
    el = heading.el.getnext()
    while el is not None and el.tag != w("sectPr"):
        if el.tag == w("p"):
            p = Para(el, MAIN, -1, None, styles)
            if p.is_heading and p.text.strip() and p.level <= heading.level:
                break
        out.append(el)
        el = el.getnext()
    return out


def insert_after(anchor, *elements):
    """Insert paragraph/table elements after a Para, block or element."""
    a = _block(anchor.el if isinstance(anchor, Para) else anchor)
    for e in elements:
        a.addnext(e)
        a = e
    return elements


def append_to_section(pkg, heading, *elements):
    """Insert elements at the end of a section (before the next heading)."""
    return insert_after(section_blocks(pkg, heading)[-1], *elements)


def delete(elements):
    """Remove paragraphs/tables (Para objects or elements). A table cell
    always keeps at least one empty paragraph."""
    for e in elements:
        el = e.el if isinstance(e, Para) else e
        parent = el.getparent()
        if parent is None:
            continue
        parent.remove(el)
        if parent.tag == w("tc") and parent.find(w("p")) is None:
            etree.SubElement(parent, w("p"))


def delete_section(pkg, heading):
    blocks = section_blocks(pkg, heading)
    delete(blocks)
    return len(blocks)


# ---- tables -------------------------------------------------------------
def tables(pkg):
    body = pkg.xml(MAIN).find(w("body"))
    return [t for t in body.iter(w("tbl")) if t.getparent().tag == w("body")]


def table_rows(tbl):
    return [[_cell_text(tc) for tc in tr.findall(w("tc"))] for tr in tbl.findall(w("tr"))]


def _cell_text(tc):
    return "\n".join("".join(t.text or "" for t in p.iter(w("t"))) for p in tc.findall(w("p")))


def find_table(pkg, contains):
    """Tables whose text contains ``contains`` (e.g. a header cell)."""
    return [t for t in tables(pkg) if any(contains in c for row in table_rows(t) for c in row)]


def cell(tbl, row, col):
    trs = tbl.findall(w("tr"))
    tcs = trs[row].findall(w("tc"))
    return tcs[col]


def set_cell_text(tbl, row, col, text):
    """Set a cell's text keeping its paragraph and run formatting."""
    tc = cell(tbl, row, col)
    ps = tc.findall(w("p"))
    for p in ps[1:]:
        tc.remove(p)
    set_paragraph_text(ps[0], text)


def _merged(tbl):
    return tbl.find(".//" + w("gridSpan")) is not None or tbl.find(".//" + w("vMerge")) is not None


def add_column(tbl, after, values):
    """Insert a column after index ``after`` (-1 = first). ``values`` has one
    text per row. Width is split from the neighbouring column."""
    if _merged(tbl):
        raise NotSupported("table has merged cells; adding a column is not supported")
    rows = tbl.findall(w("tr"))
    if len(values) != len(rows):
        raise ValueError(f"need {len(rows)} values, got {len(values)}")
    src = max(after, 0)
    grid = tbl.find(w("tblGrid"))
    if grid is not None:
        cols = grid.findall(w("gridCol"))
        gc = copy.deepcopy(cols[src])
        wv = int(cols[src].get(w("w"), "0") or 0)
        if wv:
            cols[src].set(w("w"), str(wv - wv // 2))
            gc.set(w("w"), str(wv // 2))
        (cols[src].addnext(gc) if after >= 0 else cols[0].addprevious(gc))
    for tr, val in zip(rows, values):
        tcs = tr.findall(w("tc"))
        new = copy.deepcopy(tcs[src])
        tcw = new.find(w("tcPr") + "/" + w("tcW"))
        if tcw is not None and tcw.get(w("type")) == "dxa":
            half = int(tcw.get(w("w"), "0")) // 2
            tcw.set(w("w"), str(half))
            tcs[src].find(w("tcPr") + "/" + w("tcW")).set(w("w"), str(half))
        for p in new.findall(w("p"))[1:]:
            new.remove(p)
        set_paragraph_text(new.find(w("p")), str(val))
        (tcs[src].addnext(new) if after >= 0 else tcs[0].addprevious(new))


def delete_column(tbl, col):
    if _merged(tbl):
        raise NotSupported("table has merged cells; deleting a column is not supported")
    grid = tbl.find(w("tblGrid"))
    if grid is not None:
        grid.remove(grid.findall(w("gridCol"))[col])
    for tr in tbl.findall(w("tr")):
        tr.remove(tr.findall(w("tc"))[col])


def add_row(tbl, after, values):
    """Insert a row after index ``after`` copying that row's formatting."""
    rows = tbl.findall(w("tr"))
    new = copy.deepcopy(rows[after])
    for tc, val in zip(new.findall(w("tc")), values):
        for p in tc.findall(w("p"))[1:]:
            tc.remove(p)
        set_paragraph_text(tc.find(w("p")), str(val))
    rows[after].addnext(new)
    return new


def set_cell_shading(tbl, row, col, fill):
    tc = cell(tbl, row, col)
    tcpr = ensure_first(tc, "tcPr")
    put_child(tcpr, "shd", TCPR_ORDER, {"val": "clear", "color": "auto", "fill": fill.lstrip("#").upper()})


# ---- templates ----------------------------------------------------------
def clear_body(pkg):
    """Remove all body content but keep the final sectPr (page setup,
    header/footer references) — the start of a document built on a
    user template."""
    body = pkg.xml(MAIN).find(w("body"))
    for ch in list(body):
        if ch.tag != w("sectPr"):
            body.remove(ch)


def start_from_template(template_path, dst, keep_content=False):
    """Copy a user template (.dotx or .docx) to ``dst`` as a normal .docx.
    Styles, theme, numbering, headers/footers and page setup are kept; the
    body is emptied unless keep_content=True. Then fill it with python-docx
    (``docx.Document(dst)``) using the template's own style names."""
    from .package import Package
    pkg = Package(template_path)
    if pkg.is_template():
        pkg.as_document()
    if not keep_content:
        clear_body(pkg)
    return pkg.save(dst)


def style_names(pkg, kind="paragraph"):
    s = Styles(pkg)
    return {sid: s.name(sid) for sid, el in s.by_id.items() if el.get(w("type")) == kind}
