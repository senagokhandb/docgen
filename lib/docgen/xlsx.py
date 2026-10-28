"""In-place XLSX editing on top of docgen.package.Package.

Values, formulas and cell formats are written straight into the sheet /
styles XML, so tables, conditional formats, validations, charts, images,
names and external links stay untouched. After changing values or formulas
run ``scripts/xlsx_recalc.py`` so cached results are correct again.
Row/column insertion, deletion and sorting: ``scripts/xlsx_struct.py``.

    from docgen.package import Package
    from docgen import xlsx as X
    pkg = Package(src)
    X.set_value(pkg, "Satışlar!E2", 18, expect=16)
    X.set_style(pkg, "Durum!A1:C1", font={"color": "FFFFFF"}, fill="1F3864")
    pkg.save(dst)
"""
import copy
import re

from lxml import etree

from . import DocgenError, NeedsConfirmation, NotFound, NotSupported, one, require_consent
from .package import R_NS

S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
WB = "xl/workbook.xml"
SST_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
SST_REL = R_NS + "/sharedStrings"


def s(tag):
    return f"{{{S_NS}}}{tag}"


class ValueMismatch(DocgenError):
    pass


WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
            "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
            "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
            "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
            "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
FONT_ORDER = ["b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz",
              "color", "name", "family", "charset", "scheme"]
STYLESHEET_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles",
                    "dxfs", "tableStyles", "colors", "extLst"]
BUILTIN_NUMFMT = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10,
                  "0.00E+00": 11, "mm-dd-yy": 14, "d-mmm-yy": 15, "h:mm": 20, "@": 49}


def _ordered(parent, tag, order):
    el = parent.find(s(tag))
    if el is not None:
        return el
    el = etree.Element(s(tag))
    idx = order.index(tag)
    for i, ch in enumerate(parent):
        n = etree.QName(ch).localname
        if n in order and order.index(n) > idx:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


# ---- addressing ---------------------------------------------------------
def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(n):
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def split_ref(ref):
    m = _REF.match(ref)
    if not m:
        raise ValueError(f"bad cell reference {ref!r}")
    return m.group(1).upper(), int(m.group(2))


def parse_target(target):
    """'Sheet!A1' / "'My Sheet'!A1:C3" -> (sheet, [refs...])."""
    if "!" not in target:
        raise ValueError(f"use 'Sheet!A1' form, got {target!r}")
    sheet, rng = target.rsplit("!", 1)
    sheet = sheet.strip("'")
    if ":" in rng:
        (c1, r1), (c2, r2) = (split_ref(x) for x in rng.split(":"))
        refs = [f"{col_letter(c)}{r}" for r in range(min(r1, r2), max(r1, r2) + 1)
                for c in range(min(col_index(c1), col_index(c2)), max(col_index(c1), col_index(c2)) + 1)]
    else:
        c, r = split_ref(rng)
        refs = [f"{c}{r}"]
    return sheet, refs


# ---- workbook model -----------------------------------------------------
def sheets(pkg):
    """[(name, part, state)] in workbook order."""
    wb = pkg.xml(WB)
    rmap = {r["Id"]: r["part"] for r in pkg.rels(WB)}
    return [(sh.get("name"), rmap[sh.get(f"{{{R_NS}}}id")], sh.get("state", "visible"))
            for sh in wb.find(s("sheets"))]


def sheet_part(pkg, name):
    for n, part, _ in sheets(pkg):
        if n == name:
            return part
    raise NotFound(f"sheet {name!r} not found; sheets: {[n for n, _, _ in sheets(pkg)]}")


def _sst_part(pkg):
    for r in pkg.rels(WB):
        if r["Type"] == SST_REL:
            return r["part"]
    return None


def _sst(pkg, create=False):
    part = _sst_part(pkg)
    if part is None:
        if not create:
            return None, None
        part = "xl/sharedStrings.xml"
        pkg.write(part, etree.Element(s("sst"), nsmap={None: S_NS}))
        pkg.add_rel(WB, SST_REL, "sharedStrings.xml")
        pkg.set_override(part, SST_CT)
    return part, pkg.xml(part)


def _si_text(si):
    return "".join(t.text or "" for t in si.iter(s("t")) if t.getparent().tag != s("rPh"))


def shared_strings(pkg):
    _, root = _sst(pkg)
    return [_si_text(si) for si in root.findall(s("si"))] if root is not None else []


def _rows(pkg, sheet):
    return pkg.xml(sheet_part(pkg, sheet)).find(s("sheetData"))


def _cell(pkg, sheet, ref, create=False):
    data = _rows(pkg, sheet)
    letters, rnum = split_ref(ref)
    ref = f"{letters}{rnum}"
    row = None
    for r in data.findall(s("row")):
        n = int(r.get("r"))
        if n == rnum:
            row = r
            break
        if n > rnum:
            if not create:
                return None
            row = etree.Element(s("row"), r=str(rnum))
            r.addprevious(row)
            break
    if row is None:
        if not create:
            return None
        row = etree.SubElement(data, s("row"), r=str(rnum))
    ci = col_index(letters)
    for c in row.findall(s("c")):
        cr = c.get("r")
        cc = col_index(split_ref(cr)[0]) if cr else None
        if cr == ref:
            return c
        if cc is not None and cc > ci:
            if not create:
                return None
            new = etree.Element(s("c"), r=ref)
            _default_style(pkg, sheet, row, ci, new)
            c.addprevious(new)
            return new
    if not create:
        return None
    new = etree.SubElement(row, s("c"), r=ref)
    _default_style(pkg, sheet, row, ci, new)
    return new


def _default_style(pkg, sheet, row, ci, cell):
    if row.get("customFormat") in ("1", "true") and row.get("s"):
        cell.set("s", row.get("s"))
        return
    cols = pkg.xml(sheet_part(pkg, sheet)).find(s("cols"))
    for col in (cols if cols is not None else []):
        if int(col.get("min")) <= ci <= int(col.get("max")) and col.get("style"):
            cell.set("s", col.get("style"))


def get(pkg, target):
    """(value, formula) of one cell; value is the cached value for formulas."""
    sheet, refs = parse_target(target)
    c = _cell(pkg, sheet, one(refs, "cell"))
    if c is None:
        return None, None
    f = c.find(s("f"))
    formula = ("=" + f.text) if f is not None and f.text else (None if f is None else "(shared)")
    t = c.get("t")
    if t == "inlineStr":
        return _si_text(c.find(s("is"))), formula
    v = c.findtext(s("v"))
    if v is None or (v == "" and t not in ("str", "s")):
        return None, formula
    if t == "s":
        return shared_strings(pkg)[int(v)], formula
    if t == "b":
        return v == "1", formula
    if t in ("str", "e"):
        return v, formula
    num = float(v)
    return (int(num) if num.is_integer() and "." not in v and "E" not in v.upper() else num), formula


def cells(pkg, sheet):
    """Iterate (ref, value, formula) over the stored cells of a sheet."""
    for c in _rows(pkg, sheet).iter(s("c")):
        v, f = get(pkg, f"{sheet}!{c.get('r')}")
        yield c.get("r"), v, f


def find(pkg, text, *, sheet=None, exact=False, ignore_case=False):
    """Cells whose displayed value contains (or equals) ``text``."""
    norm = (lambda x: x.casefold()) if ignore_case else (lambda x: x)
    out = []
    for name, _, _ in sheets(pkg):
        if sheet and name != sheet:
            continue
        for ref, v, f in cells(pkg, name):
            sv = "" if v is None else str(v)
            if (norm(sv) == norm(text)) if exact else (norm(text) in norm(sv)):
                out.append(f"{name}!{ref}")
    return out


def column_of(pkg, sheet, header, header_row=1):
    """Column letter whose header cell equals ``header``."""
    hits = [ref for ref, v, _ in cells(pkg, sheet)
            if split_ref(ref)[1] == header_row and v is not None and str(v).strip() == header]
    return split_ref(one(hits, f"column {header!r} in {sheet}"))[0]


def rows_where(pkg, sheet, column, value, header_row=1):
    """Row numbers whose ``column`` (letter or header text) equals value."""
    col = column if re.fullmatch(r"[A-Za-z]{1,3}", column) else column_of(pkg, sheet, column, header_row)
    return [split_ref(ref)[1] for ref, v, _ in cells(pkg, sheet)
            if split_ref(ref)[0] == col.upper() and split_ref(ref)[1] > header_row and v == value]


# ---- writing values -----------------------------------------------------
def _mark_recalc(pkg):
    wb = pkg.xml(WB)
    cp = wb.find(s("calcPr"))
    if cp is None:
        cp = etree.SubElement(wb, s("calcPr"))
        ext = wb.find(s("extLst"))
        if ext is not None:
            ext.addprevious(cp)
    cp.set("fullCalcOnLoad", "1")


def _drop_calc_chain(pkg):
    for r in pkg.rels(WB):
        if r["Type"].endswith("/calcChain"):
            pkg.remove_rel(WB, r["Id"])
            pkg.remove(r["part"])
            pkg.remove_override(r["part"])


def _expand_shared(pkg, sheet, c):
    """Before overwriting the master of a shared formula, give every cell of
    the group its own explicit formula."""
    f = c.find(s("f"))
    if f is None or f.get("t") != "shared" or not f.get("ref") or not f.text:
        return
    from openpyxl.formula.translate import Translator
    si, master = f.get("si"), c.get("r")
    for other in _rows(pkg, sheet).iter(s("c")):
        of = other.find(s("f"))
        if other is c or of is None or of.get("t") != "shared" or of.get("si") != si:
            continue
        of.text = Translator("=" + f.text, origin=master).translate_formula(other.get("r"))[1:]
        for k in ("t", "si", "ref"):
            of.attrib.pop(k, None)
    for k in ("t", "si", "ref"):
        f.attrib.pop(k, None)


def _clear(c):
    for tag in ("f", "v", "is"):
        for el in c.findall(s(tag)):
            c.remove(el)
    c.attrib.pop("t", None)


def set_value(pkg, target, value, *, expect=None, allow_formula_overwrite=False):
    """Write a constant into one cell, keeping its format.
    expect=old value -> ValueMismatch if the cell holds something else.
    Overwriting a formula needs allow_formula_overwrite=True (user --onay)."""
    if isinstance(value, str) and value.startswith("="):
        raise ValueError("value starts with '=' — use set_formula()")
    sheet, refs = parse_target(target)
    ref = one(refs, "cell")
    old, formula = get(pkg, f"{sheet}!{ref}")
    if expect is not None and old != expect:
        raise ValueMismatch(f"{sheet}!{ref} holds {old!r}, not {expect!r}; nothing written")
    c = _cell(pkg, sheet, ref, create=True)
    if formula and not allow_formula_overwrite:
        raise NeedsConfirmation(f"{sheet}!{ref} contains the formula {formula}; overwriting it with a "
                                f"constant breaks the calculation (needs --onay)")
    if formula:
        require_consent(f"overwriting the formula {formula} in {sheet}!{ref}")
        _expand_shared(pkg, sheet, c)
        _drop_calc_chain(pkg)
    _clear(c)
    if value is None:
        pass
    elif isinstance(value, bool):
        c.set("t", "b")
        etree.SubElement(c, s("v")).text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        etree.SubElement(c, s("v")).text = repr(value) if isinstance(value, float) else str(value)
    elif _sst_part(pkg) is None:
        # workbook keeps strings inline: do the same instead of adding a part
        c.set("t", "inlineStr")
        t = etree.SubElement(etree.SubElement(c, s("is")), s("t"))
        t.text = value
        if value != value.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        part, sst = _sst(pkg)
        texts = [_si_text(si) for si in sst.findall(s("si"))]
        if value in texts:
            idx = texts.index(value)
        else:
            si = etree.SubElement(sst, s("si"))
            t = etree.SubElement(si, s("t"))
            t.text = value
            if value != value.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            idx = len(texts)
            sst.set("uniqueCount", str(idx + 1))
        sst.set("count", str(int(sst.get("count", "0")) + 1))
        c.set("t", "s")
        etree.SubElement(c, s("v")).text = str(idx)
    _mark_recalc(pkg)
    return old


def set_formula(pkg, target, formula, *, allow_formula_overwrite=False):
    """Write a formula ('=SUM(B2:B3)'); its result is filled in by xlsx_recalc.py."""
    if not formula.startswith("="):
        raise ValueError("formula must start with '='")
    sheet, refs = parse_target(target)
    ref = one(refs, "cell")
    _, old = get(pkg, f"{sheet}!{ref}")
    if old and old != formula and not allow_formula_overwrite:
        raise NeedsConfirmation(f"{sheet}!{ref} already contains {old}; replacing it needs --onay")
    if old and old != formula:
        require_consent(f"replacing the formula {old} in {sheet}!{ref}")
    c = _cell(pkg, sheet, ref, create=True)
    _expand_shared(pkg, sheet, c)
    _clear(c)
    etree.SubElement(c, s("f")).text = formula[1:]
    _drop_calc_chain(pkg)
    _mark_recalc(pkg)


def formula_to_value(pkg, target, *, force=False):
    """Replace formulas in a range with their cached values (needs --onay)."""
    if not force:
        raise NeedsConfirmation(f"converting formulas in {target} to fixed values needs --onay")
    require_consent(f"converting formulas in {target} to fixed values")
    sheet, refs = parse_target(target)
    for ref in refs:
        v, f = get(pkg, f"{sheet}!{ref}")
        if f:
            set_value(pkg, f"{sheet}!{ref}", v, allow_formula_overwrite=True)


# ---- styles -------------------------------------------------------------
class _Styles:
    def __init__(self, pkg):
        self.root = pkg.xml("xl/styles.xml")

    def coll(self, tag):
        return _ordered(self.root, tag, STYLESHEET_ORDER)

    def add(self, tag, el):
        """Append (or reuse an identical) entry; returns its index."""
        coll = self.coll(tag)
        key = etree.tostring(el, method="c14n")
        for i, x in enumerate(coll):
            if etree.tostring(x, method="c14n") == key:
                return i
        coll.append(el)
        coll.set("count", str(len(coll)))
        return len(coll) - 1

    def item(self, tag, idx):
        return self.coll(tag)[idx]


def _set_font(font, spec):
    def put(tag, attrs=None):
        old = font.find(s(tag))
        if old is not None:
            font.remove(old)
        el = etree.Element(s(tag), **(attrs or {}))
        idx = FONT_ORDER.index(tag)
        for i, ch in enumerate(font):
            n = etree.QName(ch).localname
            if n in FONT_ORDER and FONT_ORDER.index(n) > idx:
                font.insert(i, el)
                return
        font.append(el)

    def drop(tag):
        for el in font.findall(s(tag)):
            font.remove(el)
    for key, tag in (("bold", "b"), ("italic", "i"), ("strike", "strike")):
        if key in spec:
            put(tag) if spec[key] else drop(tag)
    if "underline" in spec:
        u = spec["underline"]
        if u:
            put("u", {} if u is True else {"val": u})
        else:
            drop("u")
    if spec.get("size"):
        put("sz", {"val": str(spec["size"])})
    if "color" in spec:
        put("color", {"rgb": "FF" + spec["color"].lstrip("#").upper()}) if spec["color"] else drop("color")
    if spec.get("name"):
        put("name", {"val": spec["name"]})
        drop("scheme")  # a theme scheme would override the explicit name


def set_style(pkg, target, *, font=None, fill=None, number_format=None, align=None):
    """Format a cell or range without touching its other formatting.
    font={'bold':True,'italic':False,'underline':False,'strike':..,'color':'FFFFFF',
          'size':14,'name':'Arial'}  (False removes that property)
    fill='1F3864' (solid) or False (no fill); number_format='#,##0.00';
    align={'horizontal':'center','vertical':'center','wrap':True}."""
    st = _Styles(pkg)
    sheet, refs = parse_target(target)
    cache = {}
    for ref in refs:
        c = _cell(pkg, sheet, ref, create=True)
        base = int(c.get("s", "0"))
        if base not in cache:
            xf = copy.deepcopy(st.item("cellXfs", base))
            if font:
                f = copy.deepcopy(st.item("fonts", int(xf.get("fontId", "0"))))
                _set_font(f, font)
                xf.set("fontId", str(st.add("fonts", f)))
                xf.set("applyFont", "1")
            if fill is not None:
                fl = etree.Element(s("fill"))
                pf = etree.SubElement(fl, s("patternFill"), patternType="solid" if fill else "none")
                if fill:
                    etree.SubElement(pf, s("fgColor"), rgb="FF" + fill.lstrip("#").upper())
                    etree.SubElement(pf, s("bgColor"), indexed="64")
                xf.set("fillId", str(st.add("fills", fl)))
                xf.set("applyFill", "1")
            if number_format is not None:
                nid = BUILTIN_NUMFMT.get(number_format)
                if nid is None:
                    nf = st.coll("numFmts")
                    nid = next((int(x.get("numFmtId")) for x in nf if x.get("formatCode") == number_format), None)
                    if nid is None:
                        nid = max([163] + [int(x.get("numFmtId")) for x in nf]) + 1
                        etree.SubElement(nf, s("numFmt"), numFmtId=str(nid), formatCode=number_format)
                        nf.set("count", str(len(nf)))
                xf.set("numFmtId", str(nid))
                xf.set("applyNumberFormat", "1")
            if align:
                al = xf.find(s("alignment"))
                if al is None:
                    al = etree.Element(s("alignment"))
                    xf.insert(0, al)
                for k, attr in (("horizontal", "horizontal"), ("vertical", "vertical"), ("wrap", "wrapText")):
                    if k in align:
                        v = align[k]
                        if v in (None, False):
                            al.attrib.pop(attr, None)
                        else:
                            al.set(attr, "1" if v is True else str(v))
                xf.set("applyAlignment", "1")
            cache[base] = st.add("cellXfs", xf)
        c.set("s", str(cache[base]))
    return len(refs)


def add_conditional_format(pkg, target, *, operator="greaterThan", formula=None, formula2=None,
                           expression=None, fill=None, font_color=None, bold=None):
    """Persistent conditional format rule on a range, e.g.
    add_conditional_format(pkg, 'Satışlar!E2:E41', operator='greaterThan', formula='20', fill='C6EFCE').
    expression='=$B2="Başarısız"' makes an expression rule instead."""
    sheet, _ = parse_target(target)
    sqref = target.rsplit("!", 1)[1].replace("$", "")
    st = _Styles(pkg)
    dxf = etree.Element(s("dxf"))
    if font_color or bold is not None:
        f = etree.SubElement(dxf, s("font"))
        if bold is not None:
            etree.SubElement(f, s("b"), val="1" if bold else "0")
        if font_color:
            etree.SubElement(f, s("color"), rgb="FF" + font_color.lstrip("#").upper())
    if fill:
        fl = etree.SubElement(dxf, s("fill"))
        pf = etree.SubElement(fl, s("patternFill"), patternType="solid")
        etree.SubElement(pf, s("fgColor"), rgb="FF" + fill.lstrip("#").upper())
        etree.SubElement(pf, s("bgColor"), rgb="FF" + fill.lstrip("#").upper())
    dxf_id = st.add("dxfs", dxf)
    ws = pkg.xml(sheet_part(pkg, sheet))
    prio = max([0] + [int(r.get("priority", "0")) for r in ws.iter(s("cfRule"))]) + 1
    cf = etree.Element(s("conditionalFormatting"), sqref=sqref)
    if expression:
        rule = etree.SubElement(cf, s("cfRule"), type="expression", dxfId=str(dxf_id), priority=str(prio))
        etree.SubElement(rule, s("formula")).text = expression.lstrip("=")
    else:
        rule = etree.SubElement(cf, s("cfRule"), type="cellIs", dxfId=str(dxf_id), priority=str(prio),
                                operator=operator)
        etree.SubElement(rule, s("formula")).text = str(formula).lstrip("=")
        if formula2 is not None:
            etree.SubElement(rule, s("formula")).text = str(formula2).lstrip("=")
    existing = ws.findall(s("conditionalFormatting"))
    if existing:
        existing[-1].addnext(cf)
    else:
        idx = WS_ORDER.index("conditionalFormatting")
        for i, ch in enumerate(ws):
            n = etree.QName(ch).localname
            if n in WS_ORDER and WS_ORDER.index(n) > idx:
                ws.insert(i, cf)
                break
        else:
            ws.append(cf)


def set_column_width(pkg, sheet, column, width):
    ws = pkg.xml(sheet_part(pkg, sheet))
    cols = _ordered(ws, "cols", WS_ORDER)
    ci = col_index(column)
    for col in cols:
        lo, hi = int(col.get("min")), int(col.get("max"))
        if lo <= ci <= hi:
            if lo != hi:
                raise NotSupported(f"column {column} is part of a grouped width range {lo}-{hi}")
            col.set("width", str(width))
            col.set("customWidth", "1")
            return
    new = etree.Element(s("col"), min=str(ci), max=str(ci), width=str(width), customWidth="1")
    after = [c for c in cols if int(c.get("min")) < ci]
    (after[-1].addnext(new) if after else cols.insert(0, new))


# ---- formula references (used by xlsx_struct guards) --------------------
_RANGE = re.compile(r"^(?:(?:'(?P<q>(?:[^']|'')+)'|(?P<s>[^!']+))!)?\$?(?P<c1>[A-Za-z]{1,3})?\$?(?P<r1>\d+)?"
                    r"(?::\$?(?P<c2>[A-Za-z]{1,3})?\$?(?P<r2>\d+)?)?$")


def defined_names(pkg):
    dn = pkg.xml(WB).find(s("definedNames"))
    return {d.get("name"): d.text for d in (dn if dn is not None else [])}


def references(formula, default_sheet, names=None):
    """[(sheet, c1, r1, c2, r2)] referenced by a formula (names resolved)."""
    from openpyxl.formula.tokenizer import Tokenizer
    names = names or {}
    out = []
    try:
        toks = Tokenizer("=" + formula.lstrip("=")).items
    except Exception:
        return out
    for t in toks:
        if t.type != "OPERAND" or t.subtype != "RANGE":
            continue
        val = t.value
        if val in names and names[val]:
            out.extend(references(names[val], default_sheet))
            continue
        m = _RANGE.match(val)
        if not m or not (m.group("c1") or m.group("r1")):
            continue
        sh = (m.group("q") or "").replace("''", "'") or m.group("s") or default_sheet
        c1 = col_index(m.group("c1")) if m.group("c1") else None
        r1 = int(m.group("r1")) if m.group("r1") else None
        c2 = col_index(m.group("c2")) if m.group("c2") else c1
        r2 = int(m.group("r2")) if m.group("r2") else r1
        out.append((sh, c1, r1, c2, r2))
    return out


def dependents_of_rows(pkg, sheet, first, last):
    """Formulas, names, conditional formats and validations that read rows
    first..last of ``sheet`` (outside those rows)."""
    names = defined_names(pkg)
    hits = []

    def touches(refs):
        return any(sh == sheet and r1 is not None and not (r2 < first or r1 > last) for sh, _, r1, _, r2 in refs)
    for name, _, _ in sheets(pkg):
        for c in _rows(pkg, name).iter(s("c")):
            f = c.find(s("f"))
            if f is None or not f.text:
                continue
            row = split_ref(c.get("r"))[1]
            if name == sheet and first <= row <= last:
                continue
            if touches(references(f.text, name, names)):
                hits.append(f"{name}!{c.get('r')} (={f.text})")
    for n, ref in names.items():
        if ref and touches(references(ref, sheet)):
            hits.append(f"name {n} ({ref})")
    return hits


def dependents_of_cols(pkg, sheet, first, last):
    names = defined_names(pkg)
    hits = []
    for name, _, _ in sheets(pkg):
        for c in _rows(pkg, name).iter(s("c")):
            f = c.find(s("f"))
            if f is None or not f.text:
                continue
            col = col_index(split_ref(c.get("r"))[0])
            if name == sheet and first <= col <= last:
                continue
            for sh, c1, _, c2, _ in references(f.text, name, names):
                if sh == sheet and c1 is not None and not (c2 < first or c1 > last):
                    hits.append(f"{name}!{c.get('r')} (={f.text})")
                    break
    return hits


# ---- features (integrity guard) -----------------------------------------
def features(pkg):
    """Workbook features that edits must not lose."""
    out = {"sheets": [n for n, _, _ in sheets(pkg)], "names": sorted(defined_names(pkg)),
           "external_links": len([r for r in pkg.rels(WB) if r["Type"].endswith("/externalLink")])}
    for name, part, _ in sheets(pkg):
        ws = pkg.xml(part)
        rels = pkg.rels(part)
        drawings = [r["part"] for r in rels if r["Type"].endswith("/drawing")]
        charts = sum(1 for d in drawings for r in pkg.rels(d) if r["Type"].endswith("/chart"))
        images = sum(1 for d in drawings for r in pkg.rels(d) if r["Type"].endswith("/image"))
        out[name] = {
            "tables": len([r for r in rels if r["Type"].endswith("/table")]),
            "conditional_formats": len(ws.findall(s("conditionalFormatting"))),
            "validations": len(ws.findall(s("dataValidations") + "/" + s("dataValidation"))),
            "merged": len(ws.findall(s("mergeCells") + "/" + s("mergeCell"))),
            "charts": charts, "images": images,
            "hyperlinks": len(ws.findall(s("hyperlinks") + "/" + s("hyperlink"))),
            "comments": len([r for r in rels if r["Type"].endswith("/comments")]),
        }
    return out
