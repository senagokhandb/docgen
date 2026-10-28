"""Formula engine behind ``scripts/xlsx_recalc.py``: computes every formula
of a workbook with MIT-licensed engines.

IronCalc (MIT OR Apache-2.0) evaluates the workbook. Formualizer (MIT OR
Apache-2.0) answers the cells whose formula uses a function IronCalc gets
wrong or lacks (``DELEGATED``: TEXT number formats — IronCalc writes 12.5 as
"12" — and AGGREGATE). Formualizer takes IronCalc's results as its inputs,
and its answers are fed back to IronCalc so dependent cells follow.

Both engines read a temporary copy of the package (the original is never
touched) where:
- relationship targets are relative: IronCalc cannot open absolute ones
  such as ``/xl/comments/comment1.xml``;
- 3D references are expanded: ``SUM(Ay1:Ay3!B2)`` -> ``SUM('Ay1'!B2,'Ay2'!B2,'Ay3'!B2)``;
- external references (``[1]Hedef!B2``) point to hidden sheets holding the
  link's cached values, which is what Excel and LibreOffice use as well;
- NUMBERVALUE gets its spaces removed and an explicit group separator
  (IronCalc gives #VALUE! for NUMBERVALUE("12,5",",")).

LOOKUP over a computed array (``LOOKUP(2,1/(A2:A9="x"),B2:B9)``) comes out
wrong in both engines: those cells are computed by LibreOffice when it is
installed. Cells no engine can compute (functions in ``IRONCALC_MISSING``
that Formualizer cannot answer either, circular references, that LOOKUP
without LibreOffice) are returned as problems; their cached value is fed to
the cells that depend on them.

``recalc_values()`` runs the engines in a child process
(``python -m docgen.calc FILE``) so a crash in native code cannot take the
caller down.
"""
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile

from lxml import etree

from . import DocgenError
from . import xlsx as X
from .package import Package, R_NS

S = X.s
WS_REL = R_NS + "/worksheet"
WS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
# error values a cell may store (ECMA-376 ST_CellErrorValue)
STORABLE_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}
# Excel functions IronCalc 0.9.1 does not have (probe of 515 Excel functions)
IRONCALC_MISSING = {
    "AGGREGATE", "BAHTTEXT", "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER", "CUBESET",
    "CUBESETCOUNT", "CUBEVALUE", "FILTERXML", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "GETPIVOTDATA", "GROUPBY", "IMAGE", "JIS", "PHONETIC", "PIVOTBY", "RTD", "WEBSERVICE",
}
DELEGATED = {"TEXT", "AGGREGATE"}
DEFAULT_TZ = "Europe/Istanbul"
_FZ_ERRORS = {"Div": "#DIV/0!", "Na": "#N/A", "NA": "#N/A", "Name": "#NAME?", "Null": "#NULL!", "Num": "#NUM!",
              "Ref": "#REF!", "Value": "#VALUE!"}


def available():
    try:
        import ironcalc  # noqa: F401
    except ImportError:
        return False
    return True


# ---- formula text -------------------------------------------------------
def functions(formula):
    """Upper-case function names used in a formula ('_xlfn.' prefixes dropped)."""
    names = re.findall(r"([A-Za-z_][A-Za-z0-9_.]*)\(", re.sub(r'"[^"]*"', "", formula or ""))
    return {re.sub(r"^(_xlfn\.)?(_xlws\.)?", "", n, flags=re.I).upper() for n in names}


def delegated(formula, used):
    """True when Formualizer answers this cell."""
    hit = used & DELEGATED
    if hit == {"TEXT"} and re.search(r'[0#][eE][+-][0#]|"[Gg]eneral"', formula or ""):
        return False  # scientific / General formats: IronCalc's TEXT is the closer one
    return bool(hit)


def _quote(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def _map_ref(value, sheet_names, ext_sheets):
    prefix, ref = value.rsplit("!", 1)
    name = prefix[1:-1].replace("''", "'") if prefix[:1] == "'" and prefix[-1:] == "'" else prefix
    m = re.match(r"\[(\d+)\](.+)$", name)
    if m:
        hidden = ext_sheets.get((int(m.group(1)), m.group(2).lower()))
        return f"{_quote(hidden)}!{ref}" if hidden else value
    if ":" in name:
        lower = [n.lower() for n in sheet_names]
        a, b = (x.lower() for x in name.split(":", 1))
        if a in lower and b in lower:
            i, j = sorted((lower.index(a), lower.index(b)))
            return ",".join(f"{_quote(sheet_names[k])}!{ref}" for k in range(i, j + 1))
    return value


class _Call:
    def __init__(self, open_text):
        self.open, self.args = open_text, [[]]

    @property
    def name(self):
        return re.sub(r"^(_xlfn\.)?(_xlws\.)?", "", self.open[:-1], flags=re.I).upper()


def _parse(formula):
    """Formula text -> list of tokens and _Call nodes (arguments split)."""
    from openpyxl.formula.tokenizer import Tokenizer
    out = []
    stack = [("top", None, out)]
    for t in Tokenizer("=" + formula).items:
        kind, args, cur = stack[-1]
        if t.type == "FUNC" and t.subtype == "OPEN":
            call = _Call(t.value)
            cur.append(call)
            stack.append(("func", call.args, call.args[-1]))
        elif t.type == "FUNC" and t.subtype == "CLOSE":
            stack.pop()
        elif t.type == "SEP" and t.subtype == "ARG" and kind == "func":
            args.append([])
            stack[-1] = ("func", args, args[-1])
        elif t.type in ("PAREN", "ARRAY") and t.subtype == "OPEN":
            cur.append(t)
            stack.append(("inline", None, cur))
        elif t.type in ("PAREN", "ARRAY") and t.subtype == "CLOSE":
            cur.append(t)
            stack.pop()
        else:
            cur.append(t)
    return out


def _render(nodes, leaf, call):
    parts = []
    for n in nodes:
        if isinstance(n, _Call):
            args = [_render(a, leaf, call) for a in n.args]
            parts.append(call(n, args) or n.open + ",".join(args) + ")")
        else:
            parts.append(leaf(n))
    return "".join(parts)


def _numbervalue(c, args):
    """NUMBERVALUE as IronCalc gets it right: spaces removed (Excel ignores
    them) and an explicit group separator ("12,5" with "," gives #VALUE! otherwise)."""
    if c.name != "NUMBERVALUE" or not 1 <= len(args) <= 3:
        return None
    text = f'SUBSTITUTE({args[0]}," ","")'
    if len(args) == 2:
        return f'{c.open}{text},{args[1]},IF({args[1]}=",",".",","))'
    return c.open + ",".join([text] + args[1:]) + ")"


def rewrite(formula, sheet_names, ext_sheets):
    """Formula text (without '=') with 3D references expanded, external
    references mapped to their hidden sheets and NUMBERVALUE made explicit;
    None when nothing changes."""
    if not formula or ("!" not in formula and "NUMBERVALUE" not in formula.upper()):
        return None
    try:
        nodes = _parse(formula)
    except Exception:
        return None

    def leaf(t):
        if t.type == "OPERAND" and t.subtype == "RANGE" and "!" in t.value:
            return _map_ref(t.value, sheet_names, ext_sheets)
        return t.value
    new = _render(nodes, leaf, _numbervalue)
    return new if new != formula else None


def lookup_over_array(formula):
    """LOOKUP whose lookup vector is computed (LOOKUP(2,1/(A:A="x"),B:B)):
    IronCalc and Formualizer both return a wrong row for it."""
    if not formula or "LOOKUP(" not in formula.upper():
        return False
    try:
        nodes = _parse(formula)
    except Exception:
        return False
    hit = []

    def call(c, args):
        vec = [t for t in (c.args[1] if len(c.args) > 1 else []) if isinstance(t, _Call) or t.type != "WHITE-SPACE"]
        if c.name == "LOOKUP" and not (len(vec) == 1 and not isinstance(vec[0], _Call)
                                       and vec[0].subtype == "RANGE"):
            hit.append(c)
    _render(nodes, lambda t: t.value, call)
    return bool(hit)


# ---- engine copy of the package -------------------------------------------
def _cached(c, sst):
    t, v = c.get("t"), c.findtext(S("v"))
    if t == "s" and v is not None:
        return "str", sst[int(v)]
    if t == "inlineStr":
        return "str", "".join(x.text or "" for x in c.iter(S("t")))
    if v is None:
        return None
    return (t if t in ("str", "b", "e") else "n"), v


def _external_sheets(pkg, used):
    """Add a hidden sheet per external-link sheet with its cached values;
    return {(link index, sheet name lower): hidden sheet name}."""
    wb = pkg.xml(X.WB)
    refs = wb.find(S("externalReferences"))
    if refs is None:
        return {}
    rels = {r["Id"]: r.get("part") for r in pkg.rels(X.WB)}
    sheets_el = wb.find(S("sheets"))
    next_id = max(int(sh.get("sheetId", "0")) for sh in sheets_el) + 1
    out = {}
    for i, er in enumerate(refs.findall(S("externalReference")), start=1):
        part = rels.get(er.get(f"{{{R_NS}}}id"))
        if not part or not pkg.has(part):
            continue
        book = pkg.xml(part).find(S("externalBook"))
        if book is None:
            continue
        names = [n.get("val") for n in book.iter(S("sheetName"))]
        data = {int(sd.get("sheetId")): sd for sd in book.iter(S("sheetData"))}
        for j, sheet in enumerate(names):
            hidden = f"_docgen_ext{i}_{j}"
            while hidden.lower() in used:
                hidden += "_"
            used.add(hidden.lower())
            ws = etree.Element(S("worksheet"), nsmap={None: X.S_NS})
            rows = etree.SubElement(ws, S("sheetData"))
            for row in (data[j] if j in data else []):
                r = etree.SubElement(rows, S("row"), r=row.get("r"))
                for cell in row:
                    c = etree.SubElement(r, S("c"), r=cell.get("r"))
                    t, v = cell.get("t"), cell.findtext(S("v"))
                    if v is None:
                        continue
                    if t in ("s", "str", "inlineStr"):
                        c.set("t", "inlineStr")
                        etree.SubElement(etree.SubElement(c, S("is")), S("t")).text = v
                    else:
                        if t in ("b", "e"):
                            c.set("t", t)
                        etree.SubElement(c, S("v")).text = v
            wpart = pkg.next_part_name("xl/worksheets/docgen_ext{}.xml")
            pkg.write(wpart, ws)
            pkg.set_override(wpart, WS_CT)
            rid = pkg.add_rel(X.WB, WS_REL, posixpath.relpath(wpart, "xl"))
            etree.SubElement(sheets_el, S("sheet"), {"name": hidden, "sheetId": str(next_id), "state": "hidden",
                                                     f"{{{R_NS}}}id": rid})
            next_id += 1
            out[(i, sheet.lower())] = hidden
    pkg.touch(X.WB)
    return out


def _relative_rels(pkg):
    for name in pkg.names("*.rels"):
        root, hit = pkg.xml(name), False
        base = posixpath.dirname(posixpath.dirname(name))
        for r in root:
            target = r.get("Target") or ""
            if r.get("TargetMode") != "External" and target.startswith("/"):
                r.set("Target", posixpath.relpath(target.lstrip("/"), base or "."))
                hit = True
        if hit:
            pkg.touch(name)


def prepare(src, dst):
    """Write the engine copy of ``src`` to ``dst``. Returns the formula cells
    {(sheet, ref): formula text} (shared children get the master's text) and
    their cached values {(sheet, ref): (t, v) or None}."""
    pkg = Package(src)
    sheet_list = X.sheets(pkg)
    names = [n for n, _, _ in sheet_list]
    ext = _external_sheets(pkg, {n.lower() for n in names})
    sst = X.shared_strings(pkg)
    formulas, cached = {}, {}
    for name, part, _ in sheet_list:
        root, masters, children = pkg.xml(part), {}, []
        for c in root.iter(S("c")):
            f = c.find(S("f"))
            if f is None:
                continue
            key = (name, c.get("r"))
            cached[key] = _cached(c, sst)
            if f.get("t") == "shared" and not f.text:
                children.append((key, f.get("si")))
                continue
            formulas[key] = f.text or ""
            if f.get("t") == "shared":
                masters[f.get("si")] = f.text
            new = rewrite(f.text, names, ext)
            if new is not None:
                f.text = new
                pkg.touch(part)
        for key, si in children:
            formulas[key] = masters.get(si, "")
    dn = pkg.xml(X.WB).find(S("definedNames"))
    for d in (dn if dn is not None else []):
        new = rewrite(d.text, names, ext)
        if new is not None:
            d.text = new
            pkg.touch(X.WB)
    _relative_rels(pkg)
    pkg.save(dst)
    return formulas, cached


# ---- engines ------------------------------------------------------------
def _number(x):
    x = float(x)
    if x != x or x in (float("inf"), float("-inf")):
        return "e", "#NUM!"
    if x.is_integer() and abs(x) < 1e15:
        return "n", str(int(x))
    return "n", repr(x)


def _split(ref):
    letters, row = X.split_ref(ref)
    return row, X.col_index(letters)


class _IronCalc:
    def __init__(self, path, tz):
        import ironcalc as ic
        self.ic = ic
        self.model = ic.load_from_xlsx(path, "en", tz)
        self.index = {p["name"]: i for i, p in enumerate(self.model.get_worksheets_properties())}

    def evaluate(self):
        self.model.evaluate()

    def value(self, sheet, ref):
        row, col = _split(ref)
        si, CT = self.index[sheet], self.ic.CellType
        t, v = self.model.get_cell_type(si, row, col), self.model.get_cell_value(si, row, col)
        if t == CT.ErrorValue:
            return "e", str(v)
        if t == CT.LogicalValue or isinstance(v, bool):
            return "b", "1" if v else "0"
        if isinstance(v, (int, float)):
            return _number(v)
        return "str", "" if v is None else str(v)

    def put(self, sheet, ref, tv):
        """Replace a cell's formula with a constant."""
        row, col = _split(ref)
        si, (t, v) = self.index[sheet], tv
        if t == "n":
            self.model.update_cell_with_number(si, row, col, float(v))
        elif t == "b":
            self.model.update_cell_with_bool(si, row, col, v in ("1", "TRUE", True))
        elif t == "e":
            self.model.set_user_input(si, row, col, v)
        else:
            self.model.update_cell_with_text(si, row, col, v)


def _fz_value(v):
    import datetime as dt
    if isinstance(v, dict) and v.get("type") == "Error":
        return "e", _FZ_ERRORS.get(v.get("kind"), f"#{v.get('kind')}")
    if isinstance(v, bool):
        return "b", "1" if v else "0"
    if isinstance(v, (int, float)):
        return _number(v)
    if v is None:
        return "n", "0"
    if isinstance(v, str):
        return "str", v
    if isinstance(v, dt.time):
        return _number((v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400)
    if isinstance(v, (dt.date, dt.datetime)):
        d = v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
        return _number((d - dt.datetime(1899, 12, 30)).total_seconds() / 86400)
    return None


def _formualizer(path, wanted, inputs):
    """Formualizer's values for ``wanted`` cells, every cell in ``inputs``
    {(sheet, ref): (t, v)} being set as a constant first."""
    import formualizer as fz
    wb = fz.load_workbook(path)
    for (sheet, ref), (t, v) in inputs.items():
        row, col = _split(ref)
        py = float(v) if t == "n" else (v == "1") if t == "b" else fz.LiteralValue.text(v)
        wb.set_value(sheet, row, col, py)
    wb.evaluate_all()
    out = {}
    for sheet, ref in wanted:
        row, col = _split(ref)
        got = _fz_value(wb.get_value(sheet, row, col))
        if got is not None:
            out[(sheet, ref)] = got
    return out


def _libreoffice(path, wanted):
    """LibreOffice's values for ``wanted`` cells; {} when it is not installed."""
    from . import office
    if not wanted or not office.soffice():
        return {}
    pkg = Package(office.convert(path, "xlsx", tempfile.mkdtemp(prefix="docgen-calc-lo-")))
    sst = X.shared_strings(pkg)
    out = {}
    for name, part, _ in X.sheets(pkg):
        for c in pkg.xml(part).iter(S("c")):
            key = (name, c.get("r"))
            if key in wanted and _cached(c, sst):
                out[key] = _cached(c, sst)
    return out


def _problem(code, used):
    missing = sorted(used & IRONCALC_MISSING)
    if missing:
        return f"{', '.join(missing)} is not supported by the formula engine"
    if code == "#CIRC!":
        return "circular reference (or depends on one)"
    return f"the formula engine cannot calculate this formula ({code})"


def _tz():
    import ironcalc as ic
    for tz in (os.environ.get("DOCGEN_TZ"), os.environ.get("TZ")):
        if tz and tz in ic.get_all_timezones():
            return tz
    return DEFAULT_TZ


def compute(src):
    """{(sheet, ref): (t, v)} for every formula cell that could be computed
    (t: n/str/b/e, v: the <v> text), {(sheet, ref): reason} for the rest,
    and notes on how the result was obtained."""
    work = tempfile.mkdtemp(prefix="docgen-calc-")
    try:
        copy = os.path.join(work, "book.xlsx")
        formulas, cached = prepare(src, copy)
        used = {k: functions(f) for k, f in formulas.items()}
        engine = _IronCalc(copy, _tz())
        engine.evaluate()
        values = {k: engine.value(*k) for k in formulas}
        notes = []

        # cells both engines get wrong: LibreOffice when installed, otherwise reported
        suspects = {k for k in formulas if lookup_over_array(formulas[k])}
        try:
            lo = _libreoffice(copy, suspects)
        except DocgenError as e:
            lo = {}
            notes.append(str(e)[:200])
        if lo:
            notes.append(f"{len(lo)} LOOKUP cells over a computed array computed by LibreOffice")
        values.update(lo)
        unsure = suspects - set(lo)

        answered = dict(lo)
        handed = [k for k in formulas if k not in suspects and delegated(formulas[k], used[k])]
        if handed:
            inputs = {k: v for k, v in values.items() if k not in handed and k not in unsure and v[0] != "e"}
            try:
                fz = _formualizer(copy, handed, inputs)
            except Exception as e:  # noqa: BLE001 — IronCalc's own result stays
                fz = {}
                notes.append(f"Formualizer could not evaluate the workbook: {str(e)[:200]}")
            # Formualizer's error wins only over an IronCalc error (it cannot read tables: Satis[Adet] -> #NAME?)
            got = {k: v for k, v in fz.items() if v[0] != "e" or (values[k][0] == "e" and v[1] in STORABLE_ERRORS)}
            answered.update(got)
            notes.append(f"{len(got)}/{len(handed)} cells with {'/'.join(sorted(DELEGATED))} "
                         "computed by Formualizer")

        # a missing function counts even without an error (IFERROR(GETPIVOTDATA(...),0) gives 0)
        problems = {k: _problem(v[1], used[k]) for k, v in values.items()
                    if k not in answered and (used[k] & IRONCALC_MISSING or (v[0] == "e" and v[1] not in STORABLE_ERRORS))}
        problems.update({k: "LOOKUP over a computed array (LOOKUP(2,1/(...),...)) is computed wrongly by the "
                            "formula engine and LibreOffice is not installed" for k in unsure})
        feed = dict(answered)
        feed.update({k: cached[k] for k in problems if cached.get(k)})
        if feed:
            for k, tv in feed.items():
                engine.put(*k, tv)
            engine.evaluate()
            values = {k: engine.value(*k) for k in formulas}
            values.update(answered)
            for k, v in values.items():
                if k not in problems and k not in answered and v[0] == "e" and v[1] not in STORABLE_ERRORS:
                    problems[k] = _problem(v[1], used[k])
        for k in problems:
            values.pop(k, None)
        return values, problems, notes
    finally:
        shutil.rmtree(work, ignore_errors=True)


def recalc_values(path, timeout=180):
    """compute() in a child process: (values, problems, notes). The result
    comes back in a file: IronCalc prints its own warnings on stdout."""
    lib = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    env = {**os.environ, "PYTHONPATH": os.pathsep.join(p for p in (lib, os.environ.get("PYTHONPATH")) if p)}
    work = tempfile.mkdtemp(prefix="docgen-calc-out-")
    out = os.path.join(work, "result.json")
    try:
        r = subprocess.run([sys.executable, "-m", "docgen.calc", str(path), out], capture_output=True, text=True,
                           timeout=timeout, env=env)
        if r.returncode != 0 or not os.path.exists(out):
            msg = (r.stderr or r.stdout).strip().splitlines()
            raise DocgenError(f"formula engine failed: {msg[-1] if msg else f'exit {r.returncode}'}")
        with open(out, encoding="utf-8") as fh:
            data = json.load(fh)
    except subprocess.TimeoutExpired:
        raise DocgenError(f"formula engine timed out after {timeout}s") from None
    except ValueError as e:
        raise DocgenError(f"formula engine returned no result: {e}") from None
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return ({(sh, ref): (t, v) for sh, ref, t, v in data["values"]},
            {(sh, ref): why for sh, ref, why in data["problems"]}, data["notes"])


def main(argv=None):
    """python -m docgen.calc FILE [RESULT.json] — JSON to RESULT.json (or stdout)."""
    args = argv if argv is not None else sys.argv[1:]
    values, problems, notes = compute(args[0])
    data = {"values": [[sh, ref, t, v] for (sh, ref), (t, v) in values.items()],
            "problems": [[sh, ref, why] for (sh, ref), why in problems.items()], "notes": notes}
    if len(args) > 1:
        with open(args[1], "w", encoding="utf-8") as fh:
            json.dump(data, fh)
    else:
        json.dump(data, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
