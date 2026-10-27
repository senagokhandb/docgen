"""In-place PPTX editing on top of docgen.package.Package.

    from docgen.package import Package
    from docgen import pptx as P, one
    pkg = Package(src)
    sh = one(P.find_shapes(pkg, slide=2, text="Kaynak: İç analiz"), "shape")
    P.set_run_props(P.runs_of(sh), size=10, italic=True)
    P.set_geometry(sh, x=P.cm(20))
    pkg.save(dst)

Slides are addressed by their 1-based position in the deck.
"""
import copy
import posixpath
import re

from lxml import etree

from . import NeedsConfirmation, NotFound, NotSupported, one
from .package import R_NS

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
PRES = "ppt/presentation.xml"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
REL_SLIDE = R_NS + "/slide"
EMU_PER_CM = 360000


def a(tag):
    return f"{{{A_NS}}}{tag}"


def p(tag):
    return f"{{{P_NS}}}{tag}"


def cm(v):
    return int(round(v * EMU_PER_CM))


def inch(v):
    return int(round(v * 914400))


def pt(v):
    return int(round(v * 12700))


RPR_CHILDREN = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst",
                "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
                "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
FILLS = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")
SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill",
              "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]
TCPR_ORDER = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill",
              "gradFill", "blipFill", "pattFill", "grpFill", "headers", "extLst"]
ALIGN = {"left": "l", "center": "ctr", "centre": "ctr", "right": "r", "justify": "just"}


def _insert_ordered(parent, el, order):
    name = etree.QName(el).localname
    idx = order.index(name)
    for i, ch in enumerate(parent):
        n = etree.QName(ch).localname
        if n in order and order.index(n) > idx:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def _solid(color):
    f = etree.Element(a("solidFill"))
    etree.SubElement(f, a("srgbClr"), val=color.lstrip("#").upper())
    return f


def _replace_fill(parent, color, order):
    for ch in list(parent):
        if etree.QName(ch).localname in FILLS:
            parent.remove(ch)
    if color is not None:
        _insert_ordered(parent, _solid(color) if color else etree.Element(a("noFill")), order)


# ---- model --------------------------------------------------------------
def _text(el):
    return "".join(t.text or "" for t in el.iter(a("t")))


class Slide:
    def __init__(self, pkg, index, part, sld_id):
        self.pkg, self.index, self.part, self.sld_id = pkg, index, part, sld_id
        self.el = pkg.xml(part)
        self.hidden = self.el.get("show") in ("0", "false")
        self.notes_part = next((r["part"] for r in pkg.rels(part) if r["Type"].endswith("/notesSlide")), None)
        self.layout_part = next((r["part"] for r in pkg.rels(part) if r["Type"].endswith("/slideLayout")), None)
        t = next((s for s in shapes(self) if s.placeholder in ("title", "ctrTitle")), None)
        self.title = t.text if t else ""

    def __repr__(self):
        return f"<Slide {self.index} {self.title!r}{' hidden' if self.hidden else ''}>"


class Shape:
    def __init__(self, slide, el, path):
        self.slide, self.el, self.path = slide, el, path
        self.kind = etree.QName(el).localname
        nv = el.find(".//" + p("cNvPr"))
        self.id = nv.get("id") if nv is not None else None
        self.name = nv.get("name") if nv is not None else ""
        ph = el.find(".//" + p("nvPr") + "/" + p("ph"))
        self.placeholder = (ph.get("type") or "body") if ph is not None else None
        self.text = "\n".join(_text(par) for par in el.iter(a("p"))) if self.kind != "grpSp" else ""

    @property
    def xfrm(self):
        x = _xfrm(self.el, create=False)
        if x is None:
            return None
        off, ext = x.find(a("off")), x.find(a("ext"))
        return (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))

    def __repr__(self):
        return f"<Shape slide {self.slide.index} {self.kind} {self.name!r} {self.text[:40]!r}>"


def slides(pkg):
    pres = pkg.xml(PRES)
    rmap = {r["Id"]: r["part"] for r in pkg.rels(PRES)}
    lst = pres.find(p("sldIdLst"))
    return [Slide(pkg, i + 1, rmap[s.get(f"{{{R_NS}}}id")], s) for i, s in enumerate(lst if lst is not None else [])]


def slide(pkg, index=None, title=None):
    ss = slides(pkg)
    if index is not None:
        if not 1 <= index <= len(ss):
            raise NotFound(f"slide {index}: deck has {len(ss)} slides")
        return ss[index - 1]
    return one([s for s in ss if title and title.casefold() in s.title.casefold()], f"slide titled {title!r}")


def shapes(sl):
    out = []

    def walk(tree, path):
        for ch in tree:
            kind = etree.QName(ch).localname
            if kind in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                out.append(Shape(sl, ch, path))
                if kind == "grpSp":
                    walk(ch, path + (ch,))
    walk(sl.el.find(p("cSld")).find(p("spTree")), ())
    return out


def find_shapes(pkg, *, slide=None, text=None, name=None, placeholder=None, kind=None):
    """Shapes filtered by slide index, contained text, shape name,
    placeholder type ('title','body',...) or kind ('sp','pic','graphicFrame')."""
    out = []
    for sl in slides(pkg):
        if slide is not None and sl.index != slide:
            continue
        for sh in shapes(sl):
            if text is not None and text not in sh.text:
                continue
            if name is not None and name != sh.name:
                continue
            if placeholder is not None and sh.placeholder not in ((placeholder,) if placeholder != "title" else ("title", "ctrTitle")):
                continue
            if kind is not None and sh.kind != kind:
                continue
            out.append(sh)
    return out


class Match:
    def __init__(self, shape, para, start, end):
        self.shape, self.para, self.start, self.end = shape, para, start, end

    @property
    def text(self):
        return _text(self.para)[self.start:self.end]

    def __repr__(self):
        t = _text(self.para)
        return (f"slide {self.shape.slide.index} | {self.shape.name!r} | …"
                f"{t[max(0, self.start - 20):self.start]}[{self.text}]{t[self.end:self.end + 20]}…")


def find(pkg, text, *, slide=None, shape=None, whole_word=False, ignore_case=False, in_title=None):
    """Occurrences of ``text`` inside slide text (per paragraph). in_title=True
    limits to title placeholders, False excludes them."""
    pat = re.escape(text)
    if whole_word:
        pat = rf"(?<!\w){pat}(?!\w)"
    rx = re.compile(pat, re.IGNORECASE if ignore_case else 0)
    out = []
    for sh in find_shapes(pkg, slide=slide):
        if shape is not None and sh.el is not getattr(shape, "el", shape):
            continue
        is_title = sh.placeholder in ("title", "ctrTitle")
        if in_title is not None and is_title != in_title:
            continue
        for para in sh.el.iter(a("p")):
            out.extend(Match(sh, para, m.start(), m.end()) for m in rx.finditer(_text(para)))
    return out


# ---- runs ---------------------------------------------------------------
def _runs(para):
    out, pos = [], 0
    for r in para:
        if r.tag in (a("r"), a("fld")):
            t = r.findtext(a("t")) or ""
            out.append((r, pos, pos + len(t)))
            pos += len(t)
    return out


def _split(r, k):
    right = copy.deepcopy(r)
    t, rt = r.find(a("t")), right.find(a("t"))
    s = t.text or ""
    t.text, rt.text = s[:k], s[k:]
    r.addnext(right)
    return r, right


def split_runs(match):
    """Run elements covering exactly the match (split at the boundaries)."""
    out = []
    for r, s, e in _runs(match.para):
        if e <= match.start or s >= match.end or s == e:
            continue
        if r.tag == a("fld") and (s < match.start or e > match.end):
            raise NotSupported("match cuts through a field (e.g. slide number)")
        el = r
        if s < match.start:
            _, el = _split(el, match.start - s)
        if e > match.end:
            el, _ = _split(el, match.end - max(s, match.start))
        out.append(el)
    return out


def runs_of(target):
    """All runs of a Shape, a paragraph element or a table cell element."""
    el = target.el if isinstance(target, Shape) else target
    return [r for r in el.iter(a("r"), a("fld"))]


def set_run_props(runs, *, bold=None, italic=None, underline=None, strike=None, color=None,
                  size=None, font=None):
    """Character formatting. None = keep, False = switch off.
    size in pt, color 'RRGGBB', font = latin typeface."""
    for r in runs:
        rpr = r.find(a("rPr"))
        if rpr is None:
            rpr = etree.Element(a("rPr"))
            r.insert(0, rpr)
        for attr, val, on, off in (("b", bold, "1", "0"), ("i", italic, "1", "0")):
            if val is not None:
                rpr.set(attr, on if val else off)
        if underline is not None:
            rpr.set("u", ("sng" if underline is True else underline) if underline else "none")
        if strike is not None:
            rpr.set("strike", "sngStrike" if strike else "noStrike")
        if size is not None:
            rpr.set("sz", str(int(round(float(size) * 100))))
        if color is not None:
            _replace_fill(rpr, color, RPR_CHILDREN)
        if font is not None:
            for tag in ("latin", "cs"):
                old = rpr.find(a(tag))
                if old is not None:
                    rpr.remove(old)
                if font:
                    _insert_ordered(rpr, etree.Element(a(tag), typeface=font), RPR_CHILDREN)


def paragraphs_of(target):
    el = target.el if isinstance(target, Shape) else target
    return list(el.iter(a("p")))


def set_paragraph_props(paras, *, align=None, line_spacing=None, space_before=None, space_after=None):
    """align: left|center|right|justify; line_spacing multiple; spacing in pt."""
    for para in paras:
        ppr = para.find(a("pPr"))
        if ppr is None:
            ppr = etree.Element(a("pPr"))
            para.insert(0, ppr)
        if align is not None:
            ppr.set("algn", ALIGN[align])
        order = ["lnSpc", "spcBef", "spcAft"]
        for tag, val, kind in (("lnSpc", line_spacing, "pct"), ("spcBef", space_before, "pts"),
                               ("spcAft", space_after, "pts")):
            if val is None:
                continue
            old = ppr.find(a(tag))
            if old is not None:
                ppr.remove(old)
            el = etree.Element(a(tag))
            if kind == "pct":
                etree.SubElement(el, a("spcPct"), val=str(int(round(val * 100000))))
            else:
                etree.SubElement(el, a("spcPts"), val=str(int(round(val * 100))))
            pos = 0
            for i, ch in enumerate(ppr):
                n = etree.QName(ch).localname
                if n in order and order.index(n) < order.index(tag):
                    pos = i + 1
            ppr.insert(pos, el)


def replace(match, new_text):
    """Replace one match; the text keeps the first run's formatting."""
    runs = split_runs(match)
    first = runs[0]
    first.find(a("t")).text = new_text
    for r in runs[1:]:
        r.getparent().remove(r)


def set_text(target, text):
    """Replace the whole text of a shape/paragraph keeping the first run's
    formatting (one paragraph per line of ``text``)."""
    el = target.el if isinstance(target, Shape) else target
    body = el.find(".//" + p("txBody")) if el.tag != a("p") else None
    paras = [el] if body is None else body.findall(a("p"))
    proto = paras[0]
    proto_run = proto.find(a("r"))
    lines = text.split("\n")
    for extra in paras[1:]:
        extra.getparent().remove(extra)
    for ch in list(proto):
        if ch.tag not in (a("pPr"), a("endParaRPr")):
            proto.remove(ch)
    cur = proto
    for i, line in enumerate(lines):
        if i:
            nxt = copy.deepcopy(proto)
            for ch in list(nxt):
                if ch.tag not in (a("pPr"), a("endParaRPr")):
                    nxt.remove(ch)
            cur.addnext(nxt)
            cur = nxt
        r = copy.deepcopy(proto_run) if proto_run is not None else etree.Element(a("r"))
        if r.find(a("t")) is None:
            etree.SubElement(r, a("t"))
        r.find(a("t")).text = line
        end = cur.find(a("endParaRPr"))
        (end.addprevious(r) if end is not None else cur.append(r))


# ---- geometry / fill / background ---------------------------------------
def _xfrm(el, create=True):
    kind = etree.QName(el).localname
    if kind == "graphicFrame":
        x = el.find(p("xfrm"))
        return x
    holder = el.find(p("grpSpPr") if kind == "grpSp" else p("spPr"))
    if holder is None:
        return None
    x = holder.find(a("xfrm"))
    if x is None and create:
        x = etree.Element(a("xfrm"))
        holder.insert(0, x)
    return x


def _layout_xfrm(shape):
    """Geometry a placeholder inherits from its layout/master."""
    ph = shape.el.find(".//" + p("nvPr") + "/" + p("ph"))
    if ph is None or shape.slide.layout_part is None:
        return None
    pkg = shape.slide.pkg
    for part in (shape.slide.layout_part,
                 next((r["part"] for r in pkg.rels(shape.slide.layout_part) if r["Type"].endswith("/slideMaster")), None)):
        if part is None:
            continue
        for sp in pkg.xml(part).iter(p("sp")):
            lph = sp.find(".//" + p("nvPr") + "/" + p("ph"))
            if lph is None:
                continue
            if (ph.get("idx") and lph.get("idx") == ph.get("idx")) or \
               (ph.get("type") and lph.get("type") == ph.get("type")):
                x = sp.find(p("spPr") + "/" + a("xfrm"))
                if x is not None:
                    return x
    return None


def set_geometry(shape, *, x=None, y=None, w=None, h=None):
    """Move/resize a shape (EMU; use cm()/inch()/pt()). Placeholders that
    inherit their position get an explicit copy first."""
    xf = _xfrm(shape.el, create=False)
    if xf is None or xf.find(a("off")) is None:
        inherited = _layout_xfrm(shape)
        if inherited is None and shape.kind != "graphicFrame":
            raise NotSupported(f"{shape!r} has no geometry to change")
        xf = _xfrm(shape.el, create=True)
        for ch in list(xf):
            xf.remove(ch)
        for ch in inherited:
            xf.append(copy.deepcopy(ch))
    off, ext = xf.find(a("off")), xf.find(a("ext"))
    for el, attr, val in ((off, "x", x), (off, "y", y), (ext, "cx", w), (ext, "cy", h)):
        if val is not None:
            el.set(attr, str(int(val)))


def scale(shape, factor):
    """Resize around the top-left corner keeping the aspect ratio."""
    x, y, w, h = shape.xfrm or (None, None, None, None)
    if w is None:
        raise NotSupported(f"{shape!r} has no explicit size")
    set_geometry(shape, w=int(w * factor), h=int(h * factor))


def slide_size(pkg):
    s = pkg.xml(PRES).find(p("sldSz"))
    return int(s.get("cx")), int(s.get("cy"))


def set_fill(shape, color):
    """Solid fill of a shape ('RRGGBB'); False = no fill."""
    sppr = shape.el.find(p("spPr"))
    if sppr is None:
        raise NotSupported(f"{shape!r} has no shape properties")
    _replace_fill(sppr, color, SPPR_ORDER)


def set_background(sl, color):
    """Solid background for one slide only (the master is not touched)."""
    csld = sl.el.find(p("cSld"))
    old = csld.find(p("bg"))
    if old is not None:
        csld.remove(old)
    bg = etree.Element(p("bg"))
    bgpr = etree.SubElement(bg, p("bgPr"))
    bgpr.append(_solid(color))
    etree.SubElement(bgpr, a("effectLst"))
    csld.insert(0, bg)


# ---- notes --------------------------------------------------------------
def _notes_body(pkg, sl):
    if sl.notes_part is None:
        raise NotSupported(f"slide {sl.index} has no notes page")
    for sp in pkg.xml(sl.notes_part).iter(p("sp")):
        ph = sp.find(".//" + p("nvPr") + "/" + p("ph"))
        if ph is not None and ph.get("type") == "body":
            return sp
    raise NotSupported(f"slide {sl.index}: notes page has no body placeholder")


def get_notes(pkg, index):
    sl = slide(pkg, index)
    if sl.notes_part is None:
        return ""
    return "\n".join(_text(x) for x in _notes_body(pkg, sl).iter(a("p")))


def set_notes(pkg, index, text):
    body = _notes_body(pkg, slide(pkg, index))
    if body.find(".//" + a("r")) is None:
        para = body.find(".//" + a("p"))
        r = etree.Element(a("r"))
        etree.SubElement(r, a("t")).text = ""
        para.insert(1 if para.find(a("pPr")) is not None else 0, r)
    set_text(body, text)


# ---- tables -------------------------------------------------------------
def tables(pkg, slide=None):
    return [sh for sh in find_shapes(pkg, slide=slide, kind="graphicFrame") if sh.el.find(".//" + a("tbl")) is not None]


def table_rows(shape):
    tbl = shape.el.find(".//" + a("tbl"))
    return [[_text(tc) for tc in tr.findall(a("tc"))] for tr in tbl.findall(a("tr"))]


def cell(shape, row, col):
    tbl = shape.el.find(".//" + a("tbl"))
    return tbl.findall(a("tr"))[row].findall(a("tc"))[col]


def set_cell_text(shape, row, col, text):
    tc = cell(shape, row, col)
    body = tc.find(a("txBody"))
    paras = body.findall(a("p"))
    if paras[0].find(a("r")) is None:
        r = etree.SubElement(paras[0], a("r"))
        etree.SubElement(r, a("t"))
        end = paras[0].find(a("endParaRPr"))
        if end is not None:
            end.addprevious(r)
    set_text(paras[0], text)
    for extra in paras[1:]:
        body.remove(extra)


def set_cell_fill(shape, row, col, color):
    tc = cell(shape, row, col)
    tcpr = tc.find(a("tcPr"))
    if tcpr is None:
        tcpr = etree.SubElement(tc, a("tcPr"))
    _replace_fill(tcpr, color, TCPR_ORDER)


def _merged_tbl(tbl):
    return any(tc.get(k) for tc in tbl.iter(a("tc")) for k in ("gridSpan", "rowSpan", "hMerge", "vMerge"))


def add_column(shape, after, values):
    """Insert a table column after index ``after`` (width split from it)."""
    tbl = shape.el.find(".//" + a("tbl"))
    if _merged_tbl(tbl):
        raise NotSupported("table has merged cells; adding a column is not supported")
    rows = tbl.findall(a("tr"))
    if len(values) != len(rows):
        raise ValueError(f"need {len(rows)} values, got {len(values)}")
    src = max(after, 0)
    cols = tbl.find(a("tblGrid")).findall(a("gridCol"))
    wv = int(cols[src].get("w"))
    new = copy.deepcopy(cols[src])
    cols[src].set("w", str(wv - wv // 2))
    new.set("w", str(wv // 2))
    for ext in new.findall(a("extLst")):
        new.remove(ext)
    (cols[src].addnext(new) if after >= 0 else cols[0].addprevious(new))
    for tr, val in zip(rows, values):
        tcs = tr.findall(a("tc"))
        tc = copy.deepcopy(tcs[src])
        (tcs[src].addnext(tc) if after >= 0 else tcs[0].addprevious(tc))
    for i, val in enumerate(values):
        set_cell_text(shape, i, src + 1 if after >= 0 else 0, str(val))


# ---- slide operations ---------------------------------------------------
def _section_lists(pkg):
    pres = pkg.xml(PRES)
    return [s.find(f"{{{P14_NS}}}sldIdLst") for s in pres.iter(f"{{{P14_NS}}}section")]


def _update_app(pkg):
    if not pkg.has("docProps/app.xml"):
        return
    app = pkg.xml("docProps/app.xml")
    ns = app.nsmap.get(None)
    ss = slides(pkg)
    for tag, val in (("Slides", len(ss)), ("HiddenSlides", sum(s.hidden for s in ss)),
                     ("Notes", sum(1 for s in ss if s.notes_part))):
        el = app.find(f"{{{ns}}}{tag}")
        if el is not None:
            el.text = str(val)


def incoming_links(pkg, index):
    """Other slides that hyperlink to slide ``index``."""
    target = slide(pkg, index).part
    return [s.index for s in slides(pkg) if s.part != target and
            any(r.get("part") == target for r in pkg.rels(s.part))]


def delete_slide(pkg, index, force=False):
    """Delete a slide with its notes page. Refuses (NeedsConfirmation) when
    another slide links to it, unless force=True."""
    sl = slide(pkg, index)
    links = incoming_links(pkg, index)
    if links and not force:
        raise NeedsConfirmation(f"slide {index} ({sl.title!r}) is the target of links on slide(s) {links}; "
                                f"deleting it breaks them")
    pres = pkg.xml(PRES)
    rid = sl.sld_id.get(f"{{{R_NS}}}id")
    pres.find(p("sldIdLst")).remove(sl.sld_id)
    for lst in _section_lists(pkg):
        for e in list(lst):
            if e.get("id") == sl.sld_id.get("id"):
                lst.remove(e)
    pkg.remove_rel(PRES, rid)
    gone = [sl.part, pkg.rels_name(sl.part)]
    if sl.notes_part:
        gone += [sl.notes_part, pkg.rels_name(sl.notes_part)]
    for g in gone:
        if pkg.has(g):
            pkg.remove(g)
            pkg.remove_override(g)
    _update_app(pkg)


def move_slide(pkg, index, new_index):
    """Move slide ``index`` to position ``new_index`` (1-based). Sections
    keep their slides; a slide moved across a boundary joins the section of
    its new neighbour."""
    ss = slides(pkg)
    lst = pkg.xml(PRES).find(p("sldIdLst"))
    el = ss[index - 1].sld_id
    lst.remove(el)
    lst.insert(new_index - 1, el)
    secs = _section_lists(pkg)
    if secs:
        sid = el.get("id")
        owner = {e.get("id"): i for i, l in enumerate(secs) for e in l}
        order = [x.get("id") for x in lst]
        pos = order.index(sid)
        neigh = order[pos - 1] if pos > 0 else (order[1] if len(order) > 1 else None)
        owner[sid] = owner.get(neigh, owner.get(sid, 0))
        for i, l in enumerate(secs):
            tag = l[0].tag if len(l) else f"{{{P14_NS}}}sldId"
            for e in list(l):
                l.remove(e)
            for x in order:
                if owner.get(x) == i:
                    etree.SubElement(l, tag, id=x)


def duplicate_slide(pkg, index, new_index=None):
    """Copy a slide (same layout, images, links) to ``new_index`` (default:
    right after it). Slides with charts/diagrams/embedded objects are refused."""
    sl = slide(pkg, index)
    for r in pkg.rels(sl.part):
        if re.search(r"/(chart|diagram\w*|oleObject|package)$", r["Type"]):
            raise NotSupported(f"slide {index} contains {r['Type'].rsplit('/', 1)[-1]}; duplicate is not supported")
    new = pkg.next_part_name("ppt/slides/slide{}.xml")
    pkg.write(new, copy.deepcopy(sl.el))
    rels = copy.deepcopy(pkg.xml(pkg.rels_name(sl.part)))
    for r in list(rels):
        if r.get("Type").endswith("/notesSlide"):
            rels.remove(r)
    pkg.write(pkg.rels_name(new), rels)
    pkg.set_override(new, SLIDE_CT)
    rid = pkg.add_rel(PRES, REL_SLIDE, posixpath.relpath(new, "ppt"))
    lst = pkg.xml(PRES).find(p("sldIdLst"))
    sid = str(max(int(x.get("id")) for x in lst) + 1)
    e = etree.Element(p("sldId"), id=sid)
    e.set(f"{{{R_NS}}}id", rid)
    lst.insert((new_index or index + 1) - 1, e)
    secs = _section_lists(pkg)
    for l in secs:
        if any(x.get("id") == sl.sld_id.get("id") for x in l):
            etree.SubElement(l, l[0].tag, id=sid)
            move_slide(pkg, (new_index or index + 1), (new_index or index + 1))
            break
    _update_app(pkg)
    return slide(pkg, new_index or index + 1)


def set_hidden(pkg, index, hidden=True):
    el = slide(pkg, index).el
    if hidden:
        el.set("show", "0")
    elif "show" in el.attrib:
        del el.attrib["show"]
    _update_app(pkg)


# ---- templates ----------------------------------------------------------
def start_from_template(template_path, dst, keep_slides=False):
    """Copy a user template (.potx or .pptx) to ``dst`` as a .pptx with its
    masters, layouts and theme. Example slides are removed unless
    keep_slides=True. Then build slides with python-pptx:
    ``prs = pptx.Presentation(dst); prs.slide_layouts`` (template layouts)."""
    from .package import Package
    pkg = Package(template_path)
    if pkg.is_template():
        pkg.as_document()
    if not keep_slides:
        for i in range(len(slides(pkg)), 0, -1):
            delete_slide(pkg, i, force=True)
    return pkg.save(dst)
