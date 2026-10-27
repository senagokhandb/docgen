"""Content-level comparison of an original file and an edited result.

``diff(original, result)`` returns entries like
    {"kind": "changed", "where": "slide 2 | 'Kaynak' (sp 7)", "what": "format",
     "text_before": ..., "text_after": ..., "slide": 2, ...}
and ``check(entries, scope)`` splits them into allowed / violations using the
scope the agent declared (see scripts/diff_check.py).
"""
import difflib
import fnmatch
import re
import zipfile

from lxml import etree

from .package import Package, canonical

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ALWAYS_OK = ("docProps/core.xml", "docProps/app.xml")


def c14n(el):
    if el is None:
        return ""
    el = etree.fromstring(etree.tostring(el))
    for x in el.iter():
        for k in list(x.attrib):
            if "rsid" in k:
                del x.attrib[k]
    return etree.tostring(el, method="c14n").decode()


def format_changed(runs_a, runs_b):
    """True if characters present in both versions changed formatting
    (catches formatting that 'bleeds' into neighbouring text)."""
    ta, tb = "".join(r[0] for r in runs_a), "".join(r[0] for r in runs_b)
    fa = [sig for text, sig in runs_a for _ in text]
    fb = [sig for text, sig in runs_b for _ in text]
    if ta == tb:
        return fa != fb
    sm = difflib.SequenceMatcher(a=ta, b=tb, autojunk=False)
    return any(fa[i1:i2] != fb[j1:j2] for op, i1, i2, j1, j2 in sm.get_opcodes() if op == "equal")


def _short(s, n=70):
    s = (s or "").replace("\n", " / ")
    return s if len(s) <= n else s[:n - 1] + "…"


def kind_of(path):
    ext = str(path).lower().rsplit(".", 1)[-1]
    return {"docx": "docx", "dotx": "docx", "docm": "docx", "pptx": "pptx", "potx": "pptx", "pptm": "pptx",
            "xlsx": "xlsx", "xltx": "xlsx", "xlsm": "xlsx", "pdf": "pdf"}.get(ext)


# ---- package parts ------------------------------------------------------
def part_changes(a, b):
    with zipfile.ZipFile(a) as za, zipfile.ZipFile(b) as zb:
        A_ = {i.filename: za.read(i.filename) for i in za.infolist() if not i.is_dir()}
        B_ = {i.filename: zb.read(i.filename) for i in zb.infolist() if not i.is_dir()}
        dirs = [i.filename for i in zb.infolist() if i.is_dir()]
    changed = sorted(n for n in A_.keys() & B_.keys() if A_[n] != B_[n])
    semantic = [n for n in changed if canonical(n, A_[n]) != canonical(n, B_[n])]
    return {"changed": changed, "semantic": semantic, "added": sorted(B_.keys() - A_.keys()),
            "removed": sorted(A_.keys() - B_.keys()), "dir_entries": dirs}


# ---- DOCX ---------------------------------------------------------------
def _docx_model(path):
    from . import docx as D
    pkg = Package(path)
    paras = []
    for p in D.paragraphs(pkg):
        runs, last = [], None
        for r in p.runs:
            sig = c14n(r.el.find(f"{{{W}}}rPr"))
            if runs and sig == last:
                runs[-1][0] += r.text
            elif r.text:
                runs.append([r.text, sig])
                last = sig
        pics = [(e.get("cx"), e.get("cy")) for e in p.el.iter(f"{{{WP}}}extent")]
        paras.append({"part": p.part, "where": p.where(), "path": " > ".join(p.path), "text": p.text,
                      "ppr": c14n(p.el.find(f"{{{W}}}pPr")), "runs": [tuple(x) for x in runs], "pics": pics})
    styles = {}
    if pkg.has("word/styles.xml"):
        for st in pkg.xml("word/styles.xml").iter(f"{{{W}}}style"):
            styles[st.get(f"{{{W}}}styleId")] = c14n(st)
    body = pkg.xml("word/document.xml").find(f"{{{W}}}body")
    sects = [c14n(s) for s in body.iter(f"{{{W}}}sectPr")]
    tables = []
    for ti, t in enumerate(x for x in body if x.tag == f"{{{W}}}tbl"):
        cell_paras = [p for p in D.paragraphs(pkg, include_headers=False) if p.table and p.table[0] == ti]
        tables.append({"cols": len(t.findall(f"{{{W}}}tblGrid/{{{W}}}gridCol")),
                       "text": " | ".join(p.text for p in cell_paras)[:200],
                       "path": " > ".join(cell_paras[0].path) if cell_paras else ""})
    return {"paras": paras, "styles": styles, "sects": sects, "tables": tables}


def _docx_diff(a, b):
    A_, B_ = _docx_model(a), _docx_model(b)
    out = []
    for part in sorted({p["part"] for p in A_["paras"]} | {p["part"] for p in B_["paras"]}):
        pa = [p for p in A_["paras"] if p["part"] == part]
        pb = [p for p in B_["paras"] if p["part"] == part]
        sm = difflib.SequenceMatcher(a=[p["text"] for p in pa], b=[p["text"] for p in pb], autojunk=False)
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == "equal" or (op == "replace" and i2 - i1 == j2 - j1):
                for x, y in zip(pa[i1:i2], pb[j1:j2]):
                    what = []
                    if x["text"] != y["text"]:
                        what.append("text")
                    if format_changed(x["runs"], y["runs"]):
                        what.append("character format")
                    if x["ppr"] != y["ppr"]:
                        what.append("paragraph format")
                    if x["pics"] != y["pics"]:
                        what.append("images")
                    if what:
                        out.append({"kind": "changed", "what": ", ".join(what), "where": y["where"],
                                    "section": [x["path"], y["path"]], "text_before": x["text"], "text_after": y["text"]})
                continue
            for x in pa[i1:i2]:
                out.append({"kind": "deleted", "what": "paragraph", "where": x["where"], "section": [x["path"]],
                            "text_before": x["text"], "text_after": ""})
            for y in pb[j1:j2]:
                out.append({"kind": "inserted", "what": "paragraph", "where": y["where"], "section": [y["path"]],
                            "text_before": "", "text_after": y["text"]})
    for sid in sorted(A_["styles"].keys() | B_["styles"].keys()):
        if A_["styles"].get(sid) != B_["styles"].get(sid):
            out.append({"kind": "changed" if sid in A_["styles"] and sid in B_["styles"] else "added/removed",
                        "what": "style", "where": f"style {sid}", "style": sid, "text_before": "", "text_after": ""})
    if A_["sects"] != B_["sects"]:
        out.append({"kind": "changed", "what": "page setup (sectPr)", "where": "section properties",
                    "text_before": "", "text_after": ""})
    ta, tb = A_["tables"], B_["tables"]
    sm = difflib.SequenceMatcher(a=[t["text"] for t in ta], b=[t["text"] for t in tb], autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal" or (op == "replace" and i2 - i1 == j2 - j1):
            for x, y in zip(ta[i1:i2], tb[j1:j2]):
                if x["cols"] != y["cols"]:
                    out.append({"kind": "changed", "what": f"table columns {x['cols']} -> {y['cols']}",
                                "where": f"table under {y['path'] or '(no heading)'}", "section": [x["path"], y["path"]],
                                "text_before": x["text"], "text_after": y["text"]})
            continue
        for x in ta[i1:i2]:
            out.append({"kind": "deleted", "what": "table", "where": f"table under {x['path'] or '(no heading)'}",
                        "section": [x["path"]], "text_before": x["text"], "text_after": ""})
        for y in tb[j1:j2]:
            out.append({"kind": "inserted", "what": "table", "where": f"table under {y['path'] or '(no heading)'}",
                        "section": [y["path"]], "text_before": "", "text_after": y["text"]})
    return out


DOCX_MODEL_PARTS = ("word/document.xml", "word/styles.xml", "word/header*.xml", "word/footer*.xml")


# ---- PPTX ---------------------------------------------------------------
def _pptx_model(path):
    from . import pptx as PX
    pkg = Package(path)
    slides = []
    for sl in PX.slides(pkg):
        shapes = {}
        for sh in PX.shapes(sl):
            paras = []
            for para in sh.el.iter(f"{{{A}}}p"):
                runs, last = [], None
                for r in para:
                    if r.tag not in (f"{{{A}}}r", f"{{{A}}}fld"):
                        continue
                    sig = c14n(r.find(f"{{{A}}}rPr"))
                    t = r.findtext(f"{{{A}}}t") or ""
                    if runs and sig == last:
                        runs[-1][0] += t
                    else:
                        runs.append([t, sig])
                        last = sig
                paras.append((c14n(para.find(f"{{{A}}}pPr")), tuple(tuple(x) for x in runs)))
            sppr = sh.el.find(f"{{{P}}}spPr")
            geo = sh.xfrm
            look = c14n(sppr) if sppr is not None else ""
            if geo:
                look = re.sub(r"<a:xfrm.*?</a:xfrm>", "", look, flags=re.S)
            tc = [c14n(x) for x in sh.el.iter(f"{{{A}}}tcPr")]
            shapes[sh.id] = {"name": sh.name, "kind": sh.kind, "text": sh.text, "geo": geo, "look": look,
                             "paras": paras, "cells": tc}
        bg = sl.el.find(f"{{{P}}}cSld/{{{P}}}bg")
        notes = PX.get_notes(pkg, sl.index) if sl.notes_part else ""
        slides.append({"index": sl.index, "part": sl.part, "title": sl.title, "hidden": sl.hidden,
                       "layout": sl.layout_part, "bg": c14n(bg), "shapes": shapes, "notes": notes})
    return slides


def _pptx_diff(a, b):
    A_, B_ = _pptx_model(a), _pptx_model(b)
    out = []
    ia, ib = {s["part"]: s for s in A_}, {s["part"]: s for s in B_}
    order_a = [s["part"] for s in A_ if s["part"] in ib]
    order_b = [s["part"] for s in B_ if s["part"] in ia]
    if order_a != order_b:
        out.append({"kind": "changed", "what": "slide order", "where": "deck", "structure": True,
                    "text_before": " | ".join(ia[p]["title"] for p in order_a),
                    "text_after": " | ".join(ib[p]["title"] for p in order_b)})
    for p in ia.keys() - ib.keys():
        out.append({"kind": "deleted", "what": "slide", "where": f"slide {ia[p]['index']}", "slide": ia[p]["index"],
                    "structure": True, "text_before": ia[p]["title"], "text_after": ""})
    for p in ib.keys() - ia.keys():
        out.append({"kind": "inserted", "what": "slide", "where": f"new slide {ib[p]['index']}", "slide_new": ib[p]["index"],
                    "structure": True, "text_before": "", "text_after": ib[p]["title"]})
    for p in ia.keys() & ib.keys():
        x, y = ia[p], ib[p]
        loc = {"slide": x["index"], "slide_new": y["index"]}
        for key, label in (("hidden", "hidden flag"), ("layout", "layout"), ("bg", "background")):
            if x[key] != y[key]:
                out.append({"kind": "changed", "what": label, "where": f"slide {x['index']}", **loc,
                            "text_before": str(x[key])[:40], "text_after": str(y[key])[:40]})
        if x["notes"] != y["notes"]:
            out.append({"kind": "changed", "what": "speaker notes", "where": f"slide {x['index']} notes", "notes": x["index"],
                        "text_before": x["notes"], "text_after": y["notes"], **loc})
        for sid in x["shapes"].keys() | y["shapes"].keys():
            sx, sy = x["shapes"].get(sid), y["shapes"].get(sid)
            where = f"slide {x['index']} | {(sx or sy)['name']!r} (id {sid})"
            if sx is None or sy is None:
                out.append({"kind": "inserted" if sx is None else "deleted", "what": "shape", "where": where, **loc,
                            "text_before": (sx or {}).get("text", ""), "text_after": (sy or {}).get("text", "")})
                continue
            what = []
            if sx["text"] != sy["text"]:
                what.append("text")
            flat = lambda paras: [r for pr in paras for r in pr[1]]
            if format_changed(flat(sx["paras"]), flat(sy["paras"])):
                what.append("character format")
            if [pr[0] for pr in sx["paras"]] != [pr[0] for pr in sy["paras"]]:
                what.append("paragraph format")
            if sx["geo"] != sy["geo"]:
                what.append(f"position/size {sx['geo']} -> {sy['geo']}")
            if sx["look"] != sy["look"] or sx["cells"] != sy["cells"]:
                what.append("fill/outline")
            if what:
                out.append({"kind": "changed", "what": ", ".join(what), "where": where, **loc,
                            "text_before": sx["text"], "text_after": sy["text"]})
    return out


PPTX_MODEL_PARTS = ("ppt/slides/slide*.xml", "ppt/notesSlides/notesSlide*.xml", "ppt/presentation.xml")


# ---- XLSX ---------------------------------------------------------------
_BUILTIN_FMT = {"0": "General", "1": "0", "2": "0.00", "3": "#,##0", "4": "#,##0.00", "9": "0%",
                "10": "0.00%", "11": "0.00E+00", "12": "# ?/?", "13": "# ??/??", "14": "mm-dd-yy",
                "15": "d-mmm-yy", "16": "d-mmm", "17": "mmm-yy", "18": "h:mm AM/PM", "19": "h:mm:ss AM/PM",
                "20": "h:mm", "21": "h:mm:ss", "22": "m/d/yy h:mm", "37": "#,##0 ;(#,##0)",
                "38": "#,##0 ;[Red](#,##0)", "39": "#,##0.00;(#,##0.00)", "40": "#,##0.00;[Red](#,##0.00)",
                "45": "mm:ss", "46": "[h]:mm:ss", "47": "mmss.0", "48": "##0.0E+0", "49": "@"}
_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                "hlink", "folHlink"]  # SpreadsheetML theme index order (dk/lt swapped)


def _tint(rgb, tint):
    import colorsys
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    return "".join(f"{round(v * 255):02X}" for v in colorsys.hls_to_rgb(h, l, s))


class _XStyles:
    """Resolves cell formats to what is visible (theme colors/fonts -> values,
    defaults normalized), so different serializations compare equal."""

    def __init__(self, pkg):
        st = pkg.xml("xl/styles.xml") if pkg.has("xl/styles.xml") else None
        q = lambda tag: (st.find(f"{{{S}}}{tag}") if st is not None else None)
        self.xfs = list(q("cellXfs")) if q("cellXfs") is not None else []
        self.fonts = list(q("fonts")) if q("fonts") is not None else []
        self.fills = list(q("fills")) if q("fills") is not None else []
        self.borders = list(q("borders")) if q("borders") is not None else []
        self.numfmts = dict(_BUILTIN_FMT)
        self.numfmts.update({n.get("numFmtId"): n.get("formatCode") for n in (q("numFmts") if q("numFmts") is not None else [])})
        self.theme, self.theme_fonts = {}, {}
        tparts = pkg.names("xl/theme/*.xml")
        if tparts:
            th = pkg.xml(tparts[0])
            cs = th.find(f".//{{{A}}}clrScheme")
            for slot in (cs if cs is not None else []):
                c = slot[0] if len(slot) else None
                if c is not None:
                    self.theme[etree.QName(slot).localname] = (c.get("val") or c.get("lastClr") or "").upper()
            for kind in ("major", "minor"):
                lat = th.find(f".//{{{A}}}{kind}Font/{{{A}}}latin")
                if lat is not None:
                    self.theme_fonts[kind] = lat.get("typeface")
        self.cache = {}

    def color(self, el):
        if el is None or el.get("auto") in ("1", "true") or el.get("indexed") == "64":
            return "auto"
        if el.get("rgb"):
            rgb = el.get("rgb")[-6:].upper()
        elif el.get("theme") is not None:
            i = int(el.get("theme"))
            rgb = self.theme.get(_THEME_ORDER[i], f"theme{i}") if i < len(_THEME_ORDER) else f"theme{i}"
        elif el.get("indexed") is not None:
            return f"indexed{el.get('indexed')}"
        else:
            return "auto"
        if el.get("tint") and len(rgb) == 6 and not rgb.startswith("THEME"):
            rgb = _tint(rgb, float(el.get("tint")))
        return rgb

    @staticmethod
    def _on(el):
        return el is not None and el.get("val", "1") not in ("0", "false")

    def font(self, f):
        g = lambda t: f.find(f"{{{S}}}{t}")
        u = g("u")
        name = g("name").get("val") if g("name") is not None else None
        scheme = g("scheme").get("val") if g("scheme") is not None else None
        if scheme in self.theme_fonts:
            name = self.theme_fonts[scheme]
        color = self.color(g("color"))
        if color == self.theme.get("dk1", "000000") or color == "000000":
            color = "auto"
        return (self._on(g("b")), self._on(g("i")), self._on(g("strike")),
                None if u is None or u.get("val") == "none" else u.get("val", "single"),
                float(g("sz").get("val")) if g("sz") is not None else None, color, name,
                g("vertAlign").get("val") if g("vertAlign") is not None else None)

    def fill(self, f):
        pf = f.find(f"{{{S}}}patternFill")
        if pf is None:
            return c14n(f)
        pt = pf.get("patternType") or "none"
        if pt == "none":
            return "none"
        return (pt, self.color(pf.find(f"{{{S}}}fgColor")), self.color(pf.find(f"{{{S}}}bgColor")) if pt != "solid" else "")

    def border(self, b):
        out = []
        for side in ("left", "right", "top", "bottom", "diagonal"):
            e = b.find(f"{{{S}}}{side}")
            style = e.get("style") if e is not None else None
            out.append((side, style, self.color(e.find(f"{{{S}}}color")) if style else None))
        return tuple(out)

    def sig(self, idx):
        if idx not in self.cache:
            if idx >= len(self.xfs):
                self.cache[idx] = ""
            else:
                xf = self.xfs[idx]
                pick = lambda lst, k: lst[int(xf.get(k, "0"))] if lst and int(xf.get(k, "0")) < len(lst) else None
                f, fl, b = pick(self.fonts, "fontId"), pick(self.fills, "fillId"), pick(self.borders, "borderId")
                al, pr = xf.find(f"{{{S}}}alignment"), xf.find(f"{{{S}}}protection")
                align = (al.get("horizontal", "general"), al.get("vertical", "bottom"),
                         al.get("wrapText") in ("1", "true"), al.get("indent", "0"), al.get("textRotation", "0")) \
                    if al is not None else ("general", "bottom", False, "0", "0")
                prot = (pr.get("locked", "1") in ("1", "true"), pr.get("hidden", "0") in ("1", "true")) \
                    if pr is not None else (True, False)
                self.cache[idx] = repr((self.numfmts.get(xf.get("numFmtId", "0"), xf.get("numFmtId", "0")),
                                        self.font(f) if f is not None else None,
                                        self.fill(fl) if fl is not None else "none",
                                        self.border(b) if b is not None else None, align, prot))
        return self.cache[idx]


def _cf_sig(cf):
    return (cf.get("sqref"), tuple((r.get("type"), r.get("operator"), tuple(x.text for x in r.findall(f"{{{S}}}formula")))
                                   for r in cf.findall(f"{{{S}}}cfRule")))


def _same_number(a, b):
    try:
        x, y = float(a), float(b)
    except (TypeError, ValueError):
        return a == b
    return abs(x - y) <= 1e-9 * max(1.0, abs(x), abs(y))


def _xlsx_model(path):
    from . import xlsx as X
    pkg = Package(path)
    st = _XStyles(pkg)
    sst = X.shared_strings(pkg)
    sheets = {}
    for name, part, state in X.sheets(pkg):
        ws = pkg.xml(part)
        cells = {}
        for c in ws.iter(f"{{{S}}}c"):
            f = c.find(f"{{{S}}}f")
            t, v = c.get("t"), c.findtext(f"{{{S}}}v")
            if t == "s" and v is not None:
                v = sst[int(v)]
            elif t == "inlineStr":
                v = "".join(x.text or "" for x in c.iter(f"{{{S}}}t"))
            formula = None if f is None else ("=" + f.text if f.text else f"shared:{f.get('si')}")
            cells[c.get("r")] = {"value": None if formula else v, "cached": v if formula else None,
                                 "formula": formula, "style": st.sig(int(c.get("s", "0")))}
        tables = []
        for r in pkg.rels(part):
            if r["Type"].endswith("/table"):
                tables.append((pkg.xml(r["part"]).get("name"), pkg.xml(r["part"]).get("ref")))
        cols = []
        for col in ws.iter(f"{{{S}}}col"):
            for i in range(int(col.get("min")), int(col.get("max")) + 1):
                if col.get("width") or col.get("hidden") in ("1", "true"):
                    cols.append((i, round(float(col.get("width", "0")), 1), col.get("hidden") in ("1", "true")))
        sheets[name] = {"state": state, "cells": cells,
                        "cf": sorted(_cf_sig(x) for x in ws.findall(f"{{{S}}}conditionalFormatting")),
                        "dv": sorted((d.get("sqref"), d.get("type"), d.findtext(f"{{{S}}}formula1"))
                                     for d in ws.iter(f"{{{S}}}dataValidation")),
                        "merged": sorted(m.get("ref") for m in ws.iter(f"{{{S}}}mergeCell")),
                        "cols": cols, "tables": sorted(tables)}
    names = X.defined_names(pkg)
    charts = {n: canonical(n, pkg.read(n)) for n in pkg.names("xl/charts/*.xml")}
    drawings = {n: canonical(n, pkg.read(n)) for n in pkg.names("xl/drawings/*.xml")}
    ext = {n: canonical(n, pkg.read(n)) for n in pkg.names("xl/externalLinks/*.xml")}
    return {"sheets": sheets, "names": names, "charts": charts, "drawings": drawings, "ext": ext}


def _xlsx_diff(a, b):
    A_, B_ = _xlsx_model(a), _xlsx_model(b)
    out = []
    for name in A_["sheets"].keys() | B_["sheets"].keys():
        x, y = A_["sheets"].get(name), B_["sheets"].get(name)
        if x is None or y is None:
            out.append({"kind": "inserted" if x is None else "deleted", "what": "sheet", "where": name, "sheet": name,
                        "structure": True, "text_before": "", "text_after": ""})
            continue
        for ref in sorted(x["cells"].keys() | y["cells"].keys()):
            cx, cy = x["cells"].get(ref), y["cells"].get(ref)
            empty = {"value": None, "cached": None, "formula": None, "style": None}
            cx, cy = cx or dict(empty, style=(cy or {}).get("style")), cy or dict(empty, style=(cx or {}).get("style"))
            base = {"sheet": name, "cell": ref, "where": f"{name}!{ref}"}
            if cx["value"] != cy["value"] or cx["formula"] != cy["formula"]:
                out.append({"kind": "changed", "what": "value/formula", **base,
                            "text_before": str(cx["formula"] or cx["value"]), "text_after": str(cy["formula"] or cy["value"])})
            elif not _same_number(cx["cached"], cy["cached"]):
                out.append({"kind": "changed", "what": "calculated result", "recalc": True, **base,
                            "text_before": str(cx["cached"]), "text_after": str(cy["cached"])})
            if cx["style"] != cy["style"]:
                out.append({"kind": "changed", "what": "cell format", **base, "text_before": str(cx["value"]),
                            "text_after": str(cy["value"])})
        for key, label in (("cf", "conditional formatting"), ("dv", "data validation"), ("merged", "merged cells"),
                           ("cols", "column widths"), ("tables", "tables"), ("state", "sheet visibility")):
            if x[key] != y[key]:
                out.append({"kind": "changed", "what": label, "where": name, "sheet": name, "feature": key,
                            "text_before": str(x[key])[:120], "text_after": str(y[key])[:120]})
    if A_["names"] != B_["names"]:
        out.append({"kind": "changed", "what": "defined names", "where": "workbook", "feature": "names",
                    "text_before": str(A_["names"])[:150], "text_after": str(B_["names"])[:150]})
    for key, label in (("charts", "charts"), ("drawings", "drawings/images"), ("ext", "external links")):
        if A_[key] != B_[key]:
            out.append({"kind": "changed", "what": label, "where": "workbook", "feature": key,
                        "text_before": str(sorted(A_[key])), "text_after": str(sorted(B_[key]))})
    return out


XLSX_MODEL_PARTS = ("xl/worksheets/sheet*.xml", "xl/styles.xml", "xl/sharedStrings.xml", "xl/workbook.xml",
                    "xl/calcChain.xml", "xl/charts/*.xml", "xl/drawings/*.xml", "xl/externalLinks/*.xml",
                    "xl/tables/*.xml")


# ---- PDF ----------------------------------------------------------------
def _pdf_diff(a, b):
    from . import pdf as F
    A_, B_ = F.lines(a), F.lines(b)
    key = lambda ls: "\n".join(l.text for l in ls if not l.band)
    out = []
    sm = difflib.SequenceMatcher(a=[key(x) for x in A_], b=[key(x) for x in B_], autojunk=False)
    pairs, gone, new = [], [], []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal" or (op == "replace" and i2 - i1 == j2 - j1):
            pairs += list(zip(range(i1, i2), range(j1, j2)))
        else:
            gone += list(range(i1, i2))
            new += list(range(j1, j2))
    for i in list(gone):  # same page text deleted here and inserted there = moved
        j = next((j for j in new if key(B_[j]) == key(A_[i])), None)
        if j is not None:
            gone.remove(i)
            new.remove(j)
            pairs.append((i, j))
            out.append({"kind": "moved", "what": "page", "where": f"page {i + 1} -> {j + 1}", "page": i + 1,
                        "structure": True, "text_before": _short(key(A_[i])), "text_after": ""})
    for i in gone:
        out.append({"kind": "deleted", "what": "page", "where": f"page {i + 1}", "page": i + 1, "structure": True,
                    "text_before": _short(key(A_[i])), "text_after": ""})
    for j in new:
        out.append({"kind": "inserted", "what": "page", "where": f"new page {j + 1}", "page_new": j + 1,
                    "structure": True, "text_before": "", "text_after": _short(key(B_[j]))})
    for i, j in pairs:
        la, lb = A_[i], B_[j]
        sig = lambda l: (l.text, round(l.x0 * 2) / 2, round(l.top * 2) / 2, F.base_font(l.font), l.size, l.color)
        sm2 = difflib.SequenceMatcher(a=[sig(l) for l in la], b=[sig(l) for l in lb], autojunk=False)
        for op, i1, i2, j1, j2 in sm2.get_opcodes():
            if op == "equal":
                continue
            olds, news = la[i1:i2], lb[j1:j2]
            band = all(l.band for l in olds + news)
            out.append({"kind": "changed", "what": "header/footer" if band else "text/format/position",
                        "where": f"page {i + 1}", "page": i + 1, "band": band,
                        "text_before": " / ".join(l.text for l in olds), "text_after": " / ".join(l.text for l in news)})
    return out


# ---- public API ---------------------------------------------------------
DIFFS = {"docx": _docx_diff, "pptx": _pptx_diff, "xlsx": _xlsx_diff, "pdf": _pdf_diff}
MODEL_PARTS = {"docx": DOCX_MODEL_PARTS, "pptx": PPTX_MODEL_PARTS, "xlsx": XLSX_MODEL_PARTS, "pdf": ()}


def diff(original, result):
    k = kind_of(original)
    entries = DIFFS[k](original, result)
    parts = part_changes(original, result) if k != "pdf" else None
    return k, entries, parts


def _section_match(paths, want):
    want_l = [w.strip().casefold() for w in re.split(r"\s*>\s*", want) if w.strip()]
    for path in paths:
        have = [h.strip().casefold() for h in path.split(">")]
        i = 0
        for h in have:
            if i < len(want_l) and want_l[i] in h:
                i += 1
        if want_l and i == len(want_l):
            return True
    return False


def _in_range(ref, rng):
    from .xlsx import split_ref, col_index
    c, r = split_ref(ref)
    if ":" in rng:
        (c1, r1), (c2, r2) = (split_ref(x) for x in rng.split(":"))
    else:
        (c1, r1), (c2, r2) = split_ref(rng), split_ref(rng)
    return col_index(c1) <= col_index(c) <= col_index(c2) and r1 <= r <= r2


def _refs_only(formula):
    """Formula with every cell/row/column reference replaced by a marker."""
    return re.sub(r"\$?\b[A-Z]{1,3}\$?\d+\b|\$?\b\d+:\$?\d+\b|\$?\b[A-Z]{1,3}:\$?[A-Z]{1,3}\b", "#", formula.upper())


def allowed(e, scope):
    """Return the scope rule that allows entry ``e`` or None."""
    tb, ta = e.get("text_before", ""), e.get("text_after", "")
    for t in scope.get("text", []):
        if t and (t in tb or t in ta):
            return f"--allow-text {t!r}"
    for sec in scope.get("section", []):
        if _section_match(e.get("section", []), sec):
            return f"--allow-section {sec!r}"
    if e.get("style") and e["style"] in scope.get("style", []):
        return f"--allow-style {e['style']}"
    for n in scope.get("slide", []):
        if e.get("slide") == n or (e.get("slide_new") == n and e.get("slide") is None):
            return f"--allow-slide {n}"
    for n in scope.get("notes", []):
        if e.get("notes") == n:
            return f"--allow-notes {n}"
    for n in scope.get("page", []):
        if e.get("page") == n or (e.get("page_new") == n and e.get("page") is None):
            return f"--allow-page {n}"
    if e.get("band") and scope.get("band"):
        return "--allow-band"
    if e.get("structure") and scope.get("structure"):
        return "--allow-structure"
    if scope.get("structure") and e.get("what") == "value/formula" and tb.startswith("=") and ta.startswith("=") \
            and _refs_only(tb) == _refs_only(ta):
        return "--allow-structure (references shifted)"
    if e.get("recalc") and scope.get("recalc"):
        return "--allow-recalc"
    for sh in scope.get("sheet", []):
        if e.get("sheet") == sh:
            return f"--allow-sheet {sh}"
    for spec in scope.get("cell", []):
        sheet, rng = spec.rsplit("!", 1)
        if e.get("sheet") == sheet.strip("'") and e.get("cell") and _in_range(e["cell"], rng.replace("$", "")):
            return f"--allow-cell {spec}"
    for f in scope.get("feature", []):
        name, _, sheet = f.partition(":")
        if e.get("feature") == name and (not sheet or e.get("sheet") == sheet):
            return f"--allow-feature {f}"
    return None


def check(original, result, scope):
    """Full report: parts + content entries split into allowed/violations,
    plus expect-text checks."""
    k, entries, parts = diff(original, result)
    ok_entries, bad = [], []
    fmt_words = ("format", "fill", "position", "images", "background", "style", "layout")
    for e in entries:
        rule = allowed(e, scope)
        if rule and scope.get("text_only") and (any(w in e["what"] for w in fmt_words)):
            rule, e = None, dict(e, what=e["what"] + " (request was text-only)")
        if rule and scope.get("format_only") and (e["kind"] != "changed" or "text" in e["what"].split(", ")):
            rule, e = None, dict(e, what=e["what"] + " (request was format-only)")
        (ok_entries if rule else bad).append(dict(e, allowed_by=rule) if rule else e)
    part_bad = []
    if parts:
        model = MODEL_PARTS[k]
        allow_parts = list(scope.get("part", []))
        if scope.get("theme"):
            allow_parts += ["word/theme/*", "ppt/theme/*", "xl/theme/*"]
        if scope.get("structure") or k == "xlsx":
            allow_parts += ["[Content_Types].xml", "*.rels"]
        match = lambda name, g: name == g or fnmatch.fnmatchcase(name, g.replace("[", "[[]"))
        for n in parts["semantic"] + parts["added"] + parts["removed"]:
            if n in ALWAYS_OK or any(match(n, g) for g in model + tuple(allow_parts)):
                continue
            if n.endswith(".rels") and any(match(n.replace("_rels/", "").rsplit(".rels", 1)[0], g) for g in model):
                continue
            part_bad.append(n)
        if parts["dir_entries"]:
            part_bad.append(f"ZIP directory entries added: {parts['dir_entries'][:3]}")
    text_after = _all_text(result, k)
    exp_bad = [f"expected text missing: {t!r}" for t in scope.get("expect", []) if t not in text_after]
    exp_bad += [f"text still present: {t!r}" for t in scope.get("gone", []) if t in text_after]
    if scope.get("require_change") and not entries:
        exp_bad.append("no content change detected")
    return {"kind": k, "ok": not (bad or part_bad or exp_bad), "violations": bad, "part_violations": part_bad,
            "expectation_failures": exp_bad, "allowed": ok_entries, "parts": parts}


def _all_text(path, k):
    if k == "pdf":
        from . import pdf as F
        return "\n".join(F.page_texts(path))
    if k == "docx":
        from . import docx as D
        return "\n".join(p.text for p in D.paragraphs(Package(path)))
    if k == "pptx":
        m = _pptx_model(path)
        return "\n".join([sh["text"] for s in m for sh in s["shapes"].values()] + [s["notes"] for s in m])
    m = _xlsx_model(path)
    return "\n".join(str(c["value"] if c["value"] is not None else c["cached"]) for s in m["sheets"].values()
                     for c in s["cells"].values())
