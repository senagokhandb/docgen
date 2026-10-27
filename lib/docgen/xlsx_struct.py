"""In-place structural spreadsheet edits: insert/delete rows or columns and
sort a range — without re-saving the workbook through another engine.

Every reference that points at the edited sheet is rewritten the way Excel
does it: cell formulas on all sheets (incl. shared/array formulas), defined
names (print titles), conditional formats and validations (ranges and
formulas), tables, merged cells, hyperlinks, auto-filters, comments,
drawing anchors and chart series. Styles, theme and every other part stay
byte-identical. Cached results must be refreshed with xlsx_recalc.py.
"""
import re

from lxml import etree

from . import NeedsConfirmation, NotSupported
from .xlsx import (WB, _drop_calc_chain, _expand_shared, _mark_recalc, col_index, col_letter, defined_names,
                   dependents_of_cols, dependents_of_rows, s, sheet_part, sheets, split_ref)

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
VML_X = "urn:schemas-microsoft-com:office:excel"
REF_ERR = "#REF!"

_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_ROWS = re.compile(r"(\$?)(\d+):(\$?)(\d+)$")
_COLS = re.compile(r"(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$")


class Shift:
    """Insert (n > 0) or delete (n < 0) |n| rows/cols at 1-based index ``at``
    on sheet ``sheet``; axis = 'row' | 'col'."""

    def __init__(self, sheet, axis, at, n):
        self.sheet, self.axis, self.at, self.n = sheet, axis, at, n

    def one(self, i):
        """New index of position i, or None if it is deleted."""
        if self.n > 0:
            return i + self.n if i >= self.at else i
        end = self.at - self.n  # first index after the deleted block
        if i >= end:
            return i + self.n
        return None if i >= self.at else i

    def span(self, a, b):
        """New (a, b) of an inclusive span, or None if it disappears."""
        if self.n > 0:
            return (a + self.n if a >= self.at else a, b + self.n if b >= self.at else b)
        end = self.at - self.n
        na = a if a < self.at else (self.at if a < end else a + self.n)
        nb = b if b < self.at else (self.at - 1 if b < end else b + self.n)
        return None if nb < na else (na, nb)

    # -- reference text ---------------------------------------------------
    def ref(self, text):
        """Shift an A1 reference/range WITHOUT sheet prefix."""
        if ":" in text:
            m = _ROWS.match(text)
            if m:
                if self.axis != "row":
                    return text
                sp = self.span(int(m.group(2)), int(m.group(4)))
                return REF_ERR if sp is None else f"{m.group(1)}{sp[0]}:{m.group(3)}{sp[1]}"
            m = _COLS.match(text)
            if m:
                if self.axis != "col":
                    return text
                sp = self.span(col_index(m.group(2)), col_index(m.group(4)))
                return REF_ERR if sp is None else f"{m.group(1)}{col_letter(sp[0])}:{m.group(3)}{col_letter(sp[1])}"
            a, b = text.split(":", 1)
            ma, mb = _CELL.match(a), _CELL.match(b)
            if not (ma and mb):
                return text
            if self.axis == "row":
                sp = self.span(int(ma.group(4)), int(mb.group(4)))
                if sp is None:
                    return REF_ERR
                return f"{ma.group(1)}{ma.group(2)}{ma.group(3)}{sp[0]}:{mb.group(1)}{mb.group(2)}{mb.group(3)}{sp[1]}"
            sp = self.span(col_index(ma.group(2)), col_index(mb.group(2)))
            if sp is None:
                return REF_ERR
            return f"{ma.group(1)}{col_letter(sp[0])}{ma.group(3)}{ma.group(4)}:{mb.group(1)}{col_letter(sp[1])}{mb.group(3)}{mb.group(4)}"
        m = _CELL.match(text)
        if not m:
            return text
        if self.axis == "row":
            r = self.one(int(m.group(4)))
            return REF_ERR if r is None else f"{m.group(1)}{m.group(2)}{m.group(3)}{r}"
        c = self.one(col_index(m.group(2)))
        return REF_ERR if c is None else f"{m.group(1)}{col_letter(c)}{m.group(3)}{m.group(4)}"

    def sqref(self, text):
        out = [self.ref(x) for x in text.split()]
        return " ".join(x for x in out if x != REF_ERR)

    def formula(self, text, here):
        """Shift references to self.sheet inside a formula located on sheet ``here``."""
        from openpyxl.formula.tokenizer import Tokenizer
        try:
            tok = Tokenizer("=" + text)
        except Exception:
            return text
        changed = False
        for t in tok.items:
            if t.type != "OPERAND" or t.subtype != "RANGE":
                continue
            v = t.value
            if "!" in v:
                sh, _, rest = v.rpartition("!")
                name = sh.strip("'").replace("''", "'")
                if name != self.sheet:
                    continue
                new = self.ref(rest)
                nv = REF_ERR if new == REF_ERR else f"{sh}!{new}"
            else:
                if here != self.sheet or not (_CELL.match(v.split(":")[0]) or _ROWS.match(v) or _COLS.match(v)):
                    continue
                nv = self.ref(v)
            if nv != v:
                t.value = nv
                changed = True
        return tok.render()[1:] if changed else text


def _rewrite_sheet_cells(ws, sh):
    data = ws.find(s("sheetData"))
    if sh.axis == "row":
        for row in list(data):
            r = int(row.get("r"))
            nr = sh.one(r)
            if nr is None:
                data.remove(row)
                continue
            row.set("r", str(nr))
            for c in row.findall(s("c")):
                col, _ = split_ref(c.get("r"))
                c.set("r", f"{col}{nr}")
    else:
        for row in data:
            row.attrib.pop("spans", None)
            for c in list(row.findall(s("c"))):
                col, r = split_ref(c.get("r"))
                nc = sh.one(col_index(col))
                if nc is None:
                    row.remove(c)
                else:
                    c.set("r", f"{col_letter(nc)}{r}")
        cols = ws.find(s("cols"))
        for col in list(cols if cols is not None else []):
            sp = sh.span(int(col.get("min")), int(col.get("max")))
            if sp is None:
                cols.remove(col)
            else:
                col.set("min", str(sp[0]))
                col.set("max", str(sp[1]))


def _rewrite_ranges(ws, sh):
    for tag, attr in (("mergeCell", "ref"), ("hyperlink", "ref"), ("autoFilter", "ref"), ("dimension", "ref"),
                      ("selection", "sqref"), ("ignoredError", "sqref"), ("protectedRange", "sqref")):
        for el in list(ws.iter(s(tag))):
            if el.get(attr):
                new = sh.sqref(el.get(attr))
                if not new:
                    el.getparent().remove(el)
                else:
                    el.set(attr, new)
    for sel in ws.iter(s("selection")):
        if sel.get("activeCell"):
            new = sh.ref(sel.get("activeCell"))
            if new == REF_ERR:
                sel.attrib.pop("activeCell")
            else:
                sel.set("activeCell", new)
    for cf in list(ws.findall(s("conditionalFormatting"))):
        new = sh.sqref(cf.get("sqref"))
        if not new:
            ws.remove(cf)
            continue
        cf.set("sqref", new)
        for f in cf.iter(s("formula")):
            f.text = sh.formula(f.text or "", sh.sheet)
    dvs = ws.find(s("dataValidations"))
    for dv in list(dvs if dvs is not None else []):
        new = sh.sqref(dv.get("sqref"))
        if not new:
            dvs.remove(dv)
            continue
        dv.set("sqref", new)
        for tag in ("formula1", "formula2"):
            f = dv.find(s(tag))
            if f is not None and f.text:
                f.text = sh.formula(f.text, sh.sheet)
    if dvs is not None:
        dvs.set("count", str(len(dvs)))
    mc = ws.find(s("mergeCells"))
    if mc is not None:
        mc.set("count", str(len(mc)))
    brk = ws.find(s("rowBreaks" if sh.axis == "row" else "colBreaks"))
    for b in list(brk if brk is not None else []):
        i = sh.one(int(b.get("id")))
        (brk.remove(b) if i is None else b.set("id", str(i)))


def _rewrite_drawing(pkg, part, sh):
    root = pkg.xml(part)
    key = "row" if sh.axis == "row" else "col"
    for anchor in root:
        for pos in ("from", "to"):
            el = anchor.find(f"{{{XDR}}}{pos}/{{{XDR}}}{key}")
            if el is None:
                continue
            i = sh.one(int(el.text) + 1)
            if i is None:
                raise NotSupported(f"a picture/chart is anchored in the {sh.axis}s being deleted")
            el.text = str(i - 1)


def _rewrite_vml(pkg, part, sh):
    data = pkg.read(part).decode("utf-8", "replace")
    tag = "Row" if sh.axis == "row" else "Column"

    def repl(m):
        i = sh.one(int(m.group(1)) + 1)
        return m.group(0) if i is None else f"<x:{tag}>{i - 1}</x:{tag}>"
    new = re.sub(rf"<x:{tag}>(\d+)</x:{tag}>", repl, data)
    if new != data:
        pkg.write(part, new.encode("utf-8"))


def _apply(pkg, sh):
    target = sheet_part(pkg, sh.sheet)
    for name, part, _ in sheets(pkg):
        ws = pkg.xml(part)
        for c in ws.iter(s("c")):
            f = c.find(s("f"))
            if f is not None and f.text:
                f.text = sh.formula(f.text, name)
        if part == target:
            for f in ws.iter(s("f")):
                if f.get("ref"):
                    new = sh.ref(f.get("ref"))
                    if new == REF_ERR:
                        raise NotSupported("an array/shared formula range would be deleted")
                    f.set("ref", new)
            _rewrite_sheet_cells(ws, sh)
            _rewrite_ranges(ws, sh)
            for r in pkg.rels(part):
                kind = r["Type"].rsplit("/", 1)[-1]
                if kind == "table":
                    t = pkg.xml(r["part"])
                    new = sh.ref(t.get("ref"))
                    if new == REF_ERR:
                        raise NotSupported(f"table {t.get('name')} would be deleted completely")
                    t.set("ref", new)
                    af = t.find(s("autoFilter"))
                    if af is not None:
                        af.set("ref", new)
                    for cf in t.iter(s("calculatedColumnFormula")):
                        cf.text = sh.formula(cf.text or "", name)
                elif kind == "drawing":
                    _rewrite_drawing(pkg, r["part"], sh)
                elif kind == "vmlDrawing":
                    _rewrite_vml(pkg, r["part"], sh)
                elif kind == "comments":
                    cm = pkg.xml(r["part"])
                    for com in list(cm.iter(s("comment"))):
                        new = sh.ref(com.get("ref"))
                        (com.getparent().remove(com) if new == REF_ERR else com.set("ref", new))
                elif kind in ("pivotTable",):
                    raise NotSupported("sheet contains a pivot table; structural edits are not supported")
    dn = pkg.xml(WB).find(s("definedNames"))
    for d in (dn if dn is not None else []):
        if d.text:
            d.text = sh.formula(d.text, None)
    for chart in pkg.names("xl/charts/chart*.xml"):
        for f in pkg.xml(chart).iter(f"{{{C_NS}}}f"):
            if f.text:
                f.text = sh.formula(f.text, None)
    _drop_calc_chain(pkg)
    _mark_recalc(pkg)


def _expand_all_shared(pkg, sheet):
    ws = pkg.xml(sheet_part(pkg, sheet))
    for c in list(ws.iter(s("c"))):
        f = c.find(s("f"))
        if f is not None and f.get("t") == "shared" and f.get("ref") and f.text:
            _expand_shared(pkg, sheet, c)


def insert_rows(pkg, sheet, at, count=1):
    _expand_all_shared(pkg, sheet)
    _apply(pkg, Shift(sheet, "row", at, count))


def insert_cols(pkg, sheet, at, count=1):
    at = col_index(at) if isinstance(at, str) else at
    _expand_all_shared(pkg, sheet)
    _apply(pkg, Shift(sheet, "col", at, count))


def delete_rows(pkg, sheet, at, count=1, force=False):
    deps = dependents_of_rows(pkg, sheet, at, at + count - 1)
    if deps and not force:
        raise NeedsConfirmation(f"{len(deps)} reference(s) read rows {at}-{at + count - 1} of {sheet} and would "
                                f"become #REF!, e.g. {deps[:4]}")
    _expand_all_shared(pkg, sheet)
    _apply(pkg, Shift(sheet, "row", at, -count))


def delete_cols(pkg, sheet, at, count=1, force=False):
    at = col_index(at) if isinstance(at, str) else at
    deps = dependents_of_cols(pkg, sheet, at, at + count - 1)
    if deps and not force:
        raise NeedsConfirmation(f"{len(deps)} reference(s) read the columns to be deleted and would "
                                f"become #REF!, e.g. {deps[:4]}")
    _expand_all_shared(pkg, sheet)
    _apply(pkg, Shift(sheet, "col", at, -count))


# ---- sorting --------------------------------------------------------------
def _sort_key(c, sst):
    """Excel order: numbers < text < booleans < errors; blanks always last."""
    if c is None:
        return None
    t, v = c.get("t"), c.findtext(s("v"))
    if t == "inlineStr":
        v, t = "".join(x.text or "" for x in c.iter(s("t"))), "str"
    if v is None:
        return None
    if t == "s":
        return (1, sst[int(v)].casefold())
    if t in ("str",):
        return (1, v.casefold())
    if t == "b":
        return (2, v)
    if t == "e":
        return (3, v)
    return (0, float(v))


def sort_range(pkg, sheet, rng, by, descending=False, header=True, force=False):
    """Sort the rows of ``rng`` ('A1:H41') by column ``by`` ('G'). Whole rows
    of the range move together; formulas are translated to their new row."""
    from openpyxl.formula.translate import Translator
    from .xlsx import references, shared_strings
    (c1, r1), (c2, r2) = (split_ref(x) for x in rng.split(":"))
    lo, hi = col_index(c1), col_index(c2)
    first = r1 + (1 if header else 0)
    ws = pkg.xml(sheet_part(pkg, sheet))
    for m in ws.iter(s("mergeCell")):
        (a, b) = m.get("ref").split(":")
        if not (split_ref(b)[1] < first or split_ref(a)[1] > r2):
            raise NotSupported("the range contains merged cells; sorting it is not supported")
    for tag in ("hyperlink",):
        for el in ws.iter(s(tag)):
            col, row = split_ref(el.get("ref").split(":")[0])
            if first <= row <= r2 and lo <= col_index(col) <= hi:
                raise NotSupported("the range contains hyperlinks that would not move with their rows")
    _expand_all_shared(pkg, sheet)
    names = defined_names(pkg)
    data = ws.find(s("sheetData"))
    rows = {int(r.get("r")): r for r in data}
    block = {}
    for rn in range(first, r2 + 1):
        row = rows.get(rn)
        cells = [c for c in (row.findall(s("c")) if row is not None else [])
                 if lo <= col_index(split_ref(c.get("r"))[0]) <= hi]
        block[rn] = cells
        for c in cells:
            f = c.find(s("f"))
            if f is not None and f.text and not force:
                for sh_, _, rr1, _, rr2 in references(f.text, sheet, names):
                    if sh_ == sheet and rr1 is not None and (rr1 != rn or rr2 != rn) and not (rr2 < first or rr1 > r2):
                        raise NeedsConfirmation(f"{c.get('r')} (={f.text}) reads other rows of the range; "
                                                f"sorting would change its meaning")
    sst = shared_strings(pkg)
    keycol = col_index(by)
    keyed = []
    for rn, cells in block.items():
        kc = next((c for c in cells if col_index(split_ref(c.get("r"))[0]) == keycol), None)
        keyed.append((_sort_key(kc, sst), rn))
    present = sorted([k for k in keyed if k[0] is not None], key=lambda k: k[0], reverse=descending)
    order = [rn for _, rn in present] + [rn for k, rn in keyed if k is None]
    moved = {}
    for new_rn, old_rn in zip(range(first, r2 + 1), order):
        for c in block[old_rn]:
            col, _ = split_ref(c.get("r"))
            f = c.find(s("f"))
            if f is not None and f.text and new_rn != old_rn:
                f.text = Translator("=" + f.text, origin=c.get("r")).translate_formula(f"{col}{new_rn}")[1:]
            c.set("r", f"{col}{new_rn}")
            moved.setdefault(new_rn, []).append(c)
    for rn, cells in block.items():
        for c in cells:
            c.getparent().remove(c)
    for rn in range(first, r2 + 1):
        cells = moved.get(rn, [])
        if not cells:
            continue
        row = rows.get(rn)
        if row is None:
            row = etree.Element(s("row"), r=str(rn))
            later = [r for n, r in rows.items() if n > rn]
            (later[0].addprevious(row) if later else data.append(row))
            rows[rn] = row
        for c in sorted(cells, key=lambda c: col_index(split_ref(c.get("r"))[0])):
            ci = col_index(split_ref(c.get("r"))[0])
            after = [x for x in row.findall(s("c")) if col_index(split_ref(x.get("r"))[0]) > ci]
            (after[0].addprevious(c) if after else row.append(c))
    _drop_calc_chain(pkg)
    _mark_recalc(pkg)
    return order
