"""Workbook -> JSON model for the chat UI's spreadsheet preview (stdlib only).

The browser draws the grid; this module does everything that needs the file:
values formatted the way Excel shows them (number/date formats, Turkish
separators), every cell's formula (shared formulas expanded), styles resolved
to concrete colors (theme + tint, indexed palette), merges, sizes, frozen
panes, comments, data validation, simple conditional formats (cellIs, text,
color scales) and an approximation of built-in table styles.

    load(path) -> {"font": {...}, "styles": [...], "dxfs": [...], "sheets": [...], "active": 0}
"""
import colorsys
import csv
import math
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

MAX_ROWS, MAX_COLS, MAX_CELLS = 2000, 200, 120_000
MAX_PART = 150 * 1024 * 1024
DEC, THOU = ",", "."  # Turkish display, like Excel on a Turkish system

THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink"]  # SpreadsheetML theme index order (dk/lt swapped)
PALETTE = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 "
           "FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF "
           "660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF "
           "CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 "
           "003366 339966 003300 333300 993300 993366 333399 333333").split()
FMT_COLORS = {"black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#00FF00", "blue": "#0000FF",
              "yellow": "#FFFF00", "magenta": "#FF00FF", "cyan": "#00FFFF"}
BUILTIN_FMT = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: '#,##0 "₺";-#,##0 "₺"', 6: '#,##0 "₺";[Red]-#,##0 "₺"', 7: '#,##0.00 "₺";-#,##0.00 "₺"',
    8: '#,##0.00 "₺";[Red]-#,##0.00 "₺"', 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "dd.mm.yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
    20: "h:mm", 21: "h:mm:ss", 22: "dd.mm.yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 41: '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-',
    42: '_-* #,##0 "₺"_-;-* #,##0 "₺"_-;_-* "-" "₺"_-;_-@_-',
    43: '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
    44: '_-* #,##0.00 "₺"_-;-* #,##0.00 "₺"_-;_-* "-"?? "₺"_-;_-@_-',
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@"}
MONTHS = {"tr": "Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık".split(),
          "en": "January February March April May June July August September October November December".split()}
MONTHS3 = {"tr": "Oca Şub Mar Nis May Haz Tem Ağu Eyl Eki Kas Ara".split(),
           "en": "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()}
DAYS = {"tr": "Pazartesi Salı Çarşamba Perşembe Cuma Cumartesi Pazar".split(),
        "en": "Monday Tuesday Wednesday Thursday Friday Saturday Sunday".split()}
DAYS3 = {"tr": "Pzt Sal Çar Per Cum Cmt Paz".split(), "en": "Mon Tue Wed Thu Fri Sat Sun".split()}
ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"}


class ViewError(Exception):
    pass


# ---- references ---------------------------------------------------------------
def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_ref(ref):
    m = _CELL.match(ref or "")
    return (int(m.group(2)), col_index(m.group(1))) if m else None


def parse_range(rng):
    """'A1:C3' / 'B2' -> (r1, c1, r2, c2); whole rows/columns are clamped to the limits."""
    a, _, b = (rng or "").replace("$", "").partition(":")
    b = b or a
    if a.isalpha() and b.isalpha():
        return 1, col_index(a), MAX_ROWS, col_index(b)
    if a.isdigit() and b.isdigit():
        return int(a), 1, int(b), MAX_COLS
    p, q = parse_ref(a), parse_ref(b)
    if not p or not q:
        return None
    return min(p[0], q[0]), min(p[1], q[1]), max(p[0], q[0]), max(p[1], q[1])


def in_ranges(ranges, r, c):
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in ranges)


_FTOK = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])')


def shift_formula(formula, dr, dc):
    """Move the relative references of a shared formula from its master cell (Excel rule)."""
    def rep(m):
        if m.group(2) is None:  # string literal or quoted sheet name
            return m.group(0)
        cabs, col, rabs, row = m.groups()
        c = col_index(col) + (0 if cabs else dc)
        r = int(row) + (0 if rabs else dr)
        if not (1 <= c <= 16384 and 1 <= r <= 1048576):
            return "#REF!"
        return f"{cabs}{col_letter(c)}{rabs}{r}"
    return _FTOK.sub(rep, formula)


# ---- number formats -----------------------------------------------------------
def _sections(code):
    out, buf, quoted, i = [], "", False, 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            quoted = not quoted
        elif ch == "\\" and not quoted:
            buf += code[i:i + 2]
            i += 2
            continue
        if ch == ";" and not quoted:
            out.append(buf)
            buf = ""
        else:
            buf += ch
        i += 1
    return out + [buf]


def _tokens(sec):
    """-> (tokens, color, condition, lang)."""
    toks, color, cond, lang, i, n = [], None, None, None, 0, len(sec)
    while i < n:
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\" and i + 1 < n:
            toks.append(("lit", sec[i + 1]))
            i += 2
        elif ch == "_":
            toks.append(("lit", " "))
            i += 2
        elif ch == "*":
            i += 2  # repeat-to-fill character
        elif ch == "[":
            j = sec.find("]", i)
            j = n if j < 0 else j
            inner, i = sec[i + 1:j], j + 1
            low = inner.lower()
            if low in FMT_COLORS:
                color = FMT_COLORS[low]
            elif low.startswith("color") and low[5:].isdigit():
                k = int(low[5:]) - 1
                color = "#" + PALETTE[k + 8] if 0 <= k < 56 else None
            elif re.fullmatch(r"(<=|>=|<>|<|>|=)-?[\d.]+", inner):
                cond = inner
            elif inner.startswith("$"):
                sym, _, loc = inner[1:].partition("-")
                if sym:
                    toks.append(("lit", sym))
                lang = loc.upper()
            elif low in ("h", "hh", "m", "mm", "s", "ss"):
                toks.append(("elapsed", low[0]))
        elif sec[i:i + 7].lower() == "general":
            toks.append(("gen", ""))
            i += 7
        elif ch in "0#?":
            toks.append(("ph", ch))
            i += 1
        elif ch in ".,%@":
            toks.append(({".": "dot", ",": "comma", "%": "pct", "@": "at"}[ch], ch))
            i += 1
        elif ch in "Ee" and i + 1 < n and sec[i + 1] in "+-":
            toks.append(("exp", sec[i + 1]))
            i += 2
        elif sec[i:i + 5].upper() == "AM/PM":
            toks.append(("ampm", "AM/PM"))
            i += 5
        elif sec[i:i + 3].upper() == "A/P":
            toks.append(("ampm", "A/P"))
            i += 3
        elif ch.lower() in "ymdhs":
            j = i
            while j < n and sec[j].lower() == ch.lower():
                j += 1
            toks.append(("dt", sec[i:j].lower()))
            i = j
        else:
            toks.append(("lit", ch))
            i += 1
    return toks, color, cond, lang


def _is_date(toks):
    return any(t[0] in ("dt", "elapsed", "ampm") for t in toks)


def _group(digits):
    out = ""
    while len(digits) > 3:
        out = THOU + digits[-3:] + out
        digits = digits[:-3]
    return digits + out


def general(x):
    if x != x or x in (math.inf, -math.inf):
        return "#NUM!"
    if x == int(x) and abs(x) < 1e11:
        return str(int(x))
    a = abs(x)
    if a >= 1e11 or a < 1e-9:
        m, _, e = f"{x:.5E}".partition("E")
        return m.rstrip("0").rstrip(".").replace(".", DEC) + "E" + e[0] + e[1:].zfill(2)
    places = max(0, 10 - len(str(int(a))))
    s = f"{x:.{places}f}".rstrip("0").rstrip(".")
    return s.replace(".", DEC)


def _number(x, toks):
    """x >= 0; the caller adds the sign."""
    x *= 100 ** sum(1 for t in toks if t[0] == "pct")
    core = [k for k, t in enumerate(toks) if t[0] in ("ph", "dot", "exp", "gen")]
    if not core:
        return "".join(v for k, v in toks if k in ("lit", "pct"))
    first, last = core[0], core[-1]
    body, end = toks[first:last + 1], last
    while end + 1 < len(toks) and toks[end + 1][0] == "comma":  # trailing commas scale by 1000
        x /= 1000
        end += 1
    pre = "".join(v for k, v in toks[:first] if k in ("lit", "pct"))
    post = "".join(v for k, v in toks[end + 1:] if k in ("lit", "pct"))
    if any(t[0] == "gen" for t in body):
        return pre + general(x) + post
    e_at = next((k for k, t in enumerate(body) if t[0] == "exp"), None)
    mant = body if e_at is None else body[:e_at]
    d_at = next((k for k, t in enumerate(mant) if t[0] == "dot"), None)
    int_ph = [t[1] for t in (mant if d_at is None else mant[:d_at]) if t[0] == "ph"]
    dec_ph = [] if d_at is None else [t[1] for t in mant[d_at + 1:] if t[0] == "ph"]
    thousands = any(t[0] == "comma" for t in (mant if d_at is None else mant[:d_at]))
    exp = None
    if e_at is not None:
        exp_digits = sum(1 for t in body[e_at + 1:] if t[0] == "ph")
        step = max(1, len(int_ph)) if len(int_ph) > 1 else 1
        e = 0 if x == 0 else math.floor(math.log10(x))
        e = (e // step) * step if step > 1 else e
        x = x / (10 ** e)
        exp = (e, exp_digits, body[e_at][1])
    q = Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-len(dec_ph)), rounding=ROUND_HALF_UP)
    ipart, _, fpart = format(q, "f").partition(".")
    if exp and len(ipart) > max(1, len(int_ph)):  # 9.99 -> 10.0 after rounding
        e, digits, sign = exp
        exp = (e + 1, digits, sign)
        q = Decimal(repr(float(q) / 10)).quantize(Decimal(1).scaleb(-len(dec_ph)), rounding=ROUND_HALF_UP)
        ipart, _, fpart = format(q, "f").partition(".")
    need = sum(1 for p in dec_ph if p == "0")
    fpart = fpart.rstrip("0").ljust(need, "0")
    fpart += " " * sum(1 for p in dec_ph[len(fpart):] if p == "?")
    min_int = sum(1 for p in int_ph if p == "0")
    if ipart == "0" and min_int == 0:
        ipart = ""
    ipart = ipart.zfill(min_int)
    if thousands:
        ipart = _group(ipart)
    s = ipart + (DEC + fpart if d_at is not None else "")
    if exp:
        e, digits, sign = exp
        s += "E" + ("-" if e < 0 else ("+" if sign == "+" else "")) + str(abs(e)).zfill(digits)
    return pre + s + post


def _date(x, toks, date1904, lang):
    if x < 0:
        return "#" * 6
    if date1904:
        base = datetime(1904, 1, 1)
    else:
        base = datetime(1899, 12, 31) if x < 60 else datetime(1899, 12, 30)
    frac_digits = 0
    for k, t in enumerate(toks):
        if t[0] == "dot" and k > 0 and toks[k - 1][0] == "dt" and toks[k - 1][1][0] == "s":
            frac_digits = sum(1 for u in toks[k + 1:k + 4] if u == ("ph", "0"))
    ms = round(x * 86400 * 1000)
    if not frac_digits:
        ms = round(ms / 1000) * 1000
    dt = base + timedelta(milliseconds=ms)
    has_ampm = any(t[0] == "ampm" for t in toks)
    dts = [k for k, t in enumerate(toks) if t[0] == "dt"]
    out, skip = [], set()
    for k, (kind, v) in enumerate(toks):
        if k in skip:
            continue
        if kind == "dt":
            c, n = v[0], len(v)
            if c == "m":
                pos = dts.index(k)
                prev = toks[dts[pos - 1]][1][0] if pos > 0 else ""
                nxt = toks[dts[pos + 1]][1][0] if pos + 1 < len(dts) else ""
                if (prev == "h" or nxt == "s") and n <= 2:
                    out.append(f"{dt.minute:0{n}d}")
                elif n <= 2:
                    out.append(f"{dt.month:0{n}d}")
                elif n == 3:
                    out.append(MONTHS3[lang][dt.month - 1])
                elif n == 4:
                    out.append(MONTHS[lang][dt.month - 1])
                else:
                    out.append(MONTHS[lang][dt.month - 1][0])
            elif c == "y":
                out.append(f"{dt.year % 100:02d}" if n <= 2 else str(dt.year))
            elif c == "d":
                out.append(f"{dt.day:0{n}d}" if n <= 2 else (DAYS3 if n == 3 else DAYS)[lang][dt.weekday()])
            elif c == "h":
                h = (dt.hour % 12 or 12) if has_ampm else dt.hour
                out.append(f"{h:0{min(n, 2)}d}")
            elif c == "s":
                out.append(f"{dt.second:0{min(n, 2)}d}")
        elif kind == "elapsed":
            total = x * {"h": 24, "m": 1440, "s": 86400}[v]
            out.append(str(int(total)))
        elif kind == "ampm":
            am = dt.hour < 12
            out.append(("ÖÖ" if am else "ÖS") if lang == "tr" and v == "AM/PM" else
                       (v.split("/")[0] if am else v.split("/")[1]))
        elif kind == "dot" and frac_digits and k > 0 and toks[k - 1][0] == "dt" and toks[k - 1][1][0] == "s":
            out.append(DEC + f"{dt.microsecond // 1000:03d}"[:frac_digits])
            skip.update(range(k + 1, k + 1 + frac_digits))
        elif kind in ("lit", "dot", "comma", "pct"):
            out.append(v)
        elif kind == "ph":
            out.append(v)
    return "".join(out)


def _matches(x, cond):
    m = re.fullmatch(r"(<=|>=|<>|<|>|=)(-?[\d.]+)", cond)
    op, v = m.group(1), float(m.group(2))
    return {"<": x < v, "<=": x <= v, ">": x > v, ">=": x >= v, "=": x == v, "<>": x != v}[op]


def format_number(x, code, date1904=False):
    """-> (text, color|None, is_date)"""
    if not code or code.lower() == "general":
        return general(x), None, False
    parsed = [_tokens(s) for s in _sections(code)]
    signed = True
    if any(p[2] for p in parsed):
        p = next((p for p in parsed if p[2] and _matches(x, p[2])), None)
        if p is None:
            p = next((p for p in parsed if not p[2]), parsed[-1])
        elif p[2].startswith("<"):
            signed = False
    elif x < 0 and len(parsed) >= 2:
        p, signed = parsed[1], False
    elif x == 0 and len(parsed) >= 3:
        p, signed = parsed[2], False
    else:
        p = parsed[0]
    toks, color, _, lang = p
    lang = "en" if lang and re.fullmatch(r"0*409|EN(-\w+)?", lang) else "tr"  # English only when asked for
    if not toks:
        return "", color, False
    if _is_date(toks):
        return _date(x, toks, date1904, lang), color, True
    text = _number(abs(x) if x < 0 else x, toks)
    if x < 0 and signed and text.strip(" 0" + DEC + THOU + "%"):
        text = "-" + text
    return text, color, False


def format_text(s, code):
    secs = _sections(code or "")
    sec = secs[3] if len(secs) >= 4 else (secs[0] if len(secs) == 1 and "@" in secs[0] else None)
    if sec is None:
        return s
    toks = _tokens(sec)[0]
    return "".join(s if k == "at" else v for k, v in toks if k in ("lit", "at"))


# ---- styles -------------------------------------------------------------------
def _tint(rgb, tint):
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    return "".join(f"{round(v * 255):02X}" for v in colorsys.hls_to_rgb(h, l, s))


def _on(el):
    return el is not None and el.get("val", "1") not in ("0", "false")


class Styles:
    def __init__(self, root, theme):
        self.theme = theme
        q = lambda tag: list(root.find(tag)) if root is not None and root.find(tag) is not None else []
        custom = root.find("colors/indexedColors") if root is not None else None
        self.palette = [c.get("rgb", "")[-6:].upper() for c in custom] if custom is not None else PALETTE
        self.numfmts = dict(BUILTIN_FMT)
        self.numfmts.update({int(n.get("numFmtId")): n.get("formatCode") for n in q("numFmts")})
        fonts = [self.font(f) for f in q("fonts")]
        fills = [self.fill(f) for f in q("fills")]
        borders = [self.border(b) for b in q("borders")]
        self.default_font = fonts[0] if fonts else {"fn": "Calibri", "sz": 11}
        self.css, self.fmt = [], []
        for xf in q("cellXfs"):
            pick = lambda lst, key: lst[int(xf.get(key, "0"))] if int(xf.get(key, "0")) < len(lst) else {}
            st = dict(pick(fonts, "fontId"))
            st.update(pick(fills, "fillId"))
            st.update(pick(borders, "borderId"))
            al = xf.find("alignment")
            if al is not None:
                if al.get("horizontal") not in (None, "general"):
                    st["h"] = al.get("horizontal")
                if al.get("vertical") not in (None, "bottom"):
                    st["v"] = al.get("vertical")
                if al.get("wrapText") in ("1", "true"):
                    st["w"] = 1
                if al.get("indent") not in (None, "0"):
                    st["ind"] = int(al.get("indent"))
                if al.get("textRotation") not in (None, "0"):
                    st["rot"] = int(al.get("textRotation"))
            self.css.append(st)
            self.fmt.append(self.numfmts.get(int(xf.get("numFmtId", "0")), "General"))
        if not self.css:
            self.css, self.fmt = [dict(self.default_font)], ["General"]
        self.dxfs = []
        for d in q("dxfs"):
            st = self.font(d.find("font"), dxf=True) if d.find("font") is not None else {}
            if d.find("fill") is not None:
                st.update(self.fill(d.find("fill"), dxf=True))
            if d.find("border") is not None:
                st.update(self.border(d.find("border")))
            self.dxfs.append(st)

    def color(self, el):
        if el is None or el.get("auto") in ("1", "true"):
            return None
        rgb = None
        if el.get("rgb"):
            rgb = el.get("rgb")[-6:].upper()
        elif el.get("theme") is not None:
            i = int(el.get("theme"))
            rgb = self.theme["colors"].get(THEME_ORDER[i]) if i < len(THEME_ORDER) else None
        elif el.get("indexed") is not None:
            i = int(el.get("indexed"))
            rgb = self.palette[i] if i < len(self.palette) and i < 64 else None
        if not rgb or not re.fullmatch(r"[0-9A-F]{6}", rgb):
            return None
        if el.get("tint"):
            rgb = _tint(rgb, float(el.get("tint")))
        return "#" + rgb

    def font(self, f, dxf=False):
        g = lambda t: f.find(t)
        st = {}
        name = g("name").get("val") if g("name") is not None else None
        scheme = g("scheme").get("val") if g("scheme") is not None else None
        if scheme in self.theme["fonts"]:
            name = self.theme["fonts"][scheme]
        if name and not dxf:
            st["fn"] = name
        if g("sz") is not None and not dxf:
            st["sz"] = float(g("sz").get("val"))
        for key, tag in (("b", "b"), ("i", "i"), ("s", "strike")):
            if _on(g(tag)):
                st[key] = 1
        if g("u") is not None and g("u").get("val") != "none":
            st["u"] = g("u").get("val", "single")
        if g("vertAlign") is not None and g("vertAlign").get("val") in ("superscript", "subscript"):
            st["va"] = g("vertAlign").get("val")
        c = self.color(g("color"))
        if c:
            st["c"] = c
        return st

    def fill(self, f, dxf=False):
        pf = f.find("patternFill")
        if pf is not None:
            pt = pf.get("patternType") or ("solid" if dxf else "none")
            if pt == "none":
                return {}
            fg, bg = self.color(pf.find("fgColor")), self.color(pf.find("bgColor"))
            c = (bg or fg) if dxf else (fg or (bg if pt != "solid" else None))
            return {"bg": c} if c else {}
        gf = f.find("gradientFill")
        if gf is not None:
            stop = gf.find("stop/color")
            c = self.color(stop)
            return {"bg": c} if c else {}
        return {}

    def border(self, b):
        out = {}
        for side, key in (("left", "l"), ("right", "r"), ("top", "t"), ("bottom", "b")):
            e = b.find(side) if b.find(side) is not None else b.find({"left": "start", "right": "end"}.get(side, side))
            if e is not None and e.get("style") not in (None, "none"):
                out[key] = [e.get("style"), self.color(e.find("color")) or "#000000"]
        return {"bd": out} if out else {}


# ---- package ------------------------------------------------------------------
def _xml(z, name):
    if z.getinfo(name).file_size > MAX_PART:
        raise ViewError(f"{name} çok büyük")
    root = ET.fromstring(z.read(name))
    for el in root.iter():
        if isinstance(el.tag, str) and "}" in el.tag:
            el.tag = el.tag.split("}", 1)[1]
        for k in [k for k in el.attrib if "}" in k]:
            el.attrib[k.split("}", 1)[1]] = el.attrib.pop(k)
    return root


def _rels(z, names, part):
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if rp not in names:
        return []
    out = []
    for r in _xml(z, rp):
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target", "")
        path = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out.append({"id": r.get("Id"), "type": r.get("Type", "").rsplit("/", 1)[-1], "part": path})
    return out


def _theme(z, part):
    colors, fonts = {}, {}
    if part:
        root = _xml(z, part)
        cs = root.find(".//clrScheme")
        for slot in (cs if cs is not None else []):
            c = slot[0] if len(slot) else None
            if c is not None:
                colors[slot.tag] = (c.get("val") or c.get("lastClr") or "").upper()
        for kind in ("major", "minor"):
            lat = root.find(f".//{kind}Font/latin")
            if lat is not None and lat.get("typeface"):
                fonts[kind] = lat.get("typeface")
    return {"colors": colors, "fonts": fonts}


def _text(el):
    """Plain text of a string item (<si>, <is>, comment <text>); phonetic runs are skipped."""
    if el is None:
        return ""
    runs = el.findall("r")
    if runs:
        return "".join(r.findtext("t") or "" for r in runs)
    return el.findtext("t") or ""


# ---- table styles (approximation of Excel's built-in TableStyleLight/Medium/DarkN) --
def _table_look(name, theme):
    m = re.fullmatch(r"TableStyle(Light|Medium|Dark)(\d+)", name or "")
    if not m:
        return None
    kind, n = m.group(1), int(m.group(2))
    k, group = (n - 1) % 7, (n - 1) // 7
    base = (theme["colors"].get("dk1") or "000000") if k == 0 else (theme["colors"].get(f"accent{k}") or "4472C4")
    tint = lambda t: "#" + _tint(base, t)
    line = ["thin", tint(0.4 if k else 0.5)]
    hdr_fill = {"bg": "#" + base, "c": "#FFFFFF", "b": 1}
    if kind == "Light":
        if group == 1:
            return {"hdr": hdr_fill, "band": {}, "body": {"bd": {"b": ["thin", "#" + base]}}}
        body = {"bd": {"b": line, "r": line, "l": line, "t": line}} if group == 2 else {}
        return {"hdr": {"b": 1, "bd": {"b": ["thin", "#" + base]}}, "band": {"bg": tint(0.8 if k else 0.85)}, "body": body}
    if kind == "Medium":
        if group == 1:
            return {"hdr": hdr_fill, "body": {"bg": tint(0.8)}, "band": {"bg": tint(0.6)}}
        if group == 3:
            return {"hdr": {"b": 1, "bg": tint(0.6)}, "body": {"bg": tint(0.8)}, "band": {"bg": tint(0.6)}}
        return {"hdr": hdr_fill, "band": {"bg": tint(0.8 if k else 0.85)}, "body": {"bd": {"b": line}}}
    return {"hdr": {"bg": "#000000", "c": "#FFFFFF", "b": 1}, "body": {"bg": "#" + _tint(base, -0.25), "c": "#FFFFFF"},
            "band": {"bg": "#" + base, "c": "#FFFFFF"}}


# ---- conditional formats ------------------------------------------------------------
def _operand(text, cells):
    t = (text or "").strip()
    if t.startswith('"') and t.endswith('"'):
        return t[1:-1].replace('""', '"')
    try:
        return float(t)
    except ValueError:
        pass
    ref = parse_ref(t)
    if ref and ref in cells:
        return cells[ref].get("num", cells[ref].get("str"))
    raise ValueError(t)


def _compare(a, op, b, b2=None):
    if isinstance(a, str) or isinstance(b, str):
        a, b = str(a).lower(), str(b).lower()
        b2 = None if b2 is None else str(b2).lower()
    tests = {"equal": lambda: a == b, "notEqual": lambda: a != b, "greaterThan": lambda: a > b,
             "lessThan": lambda: a < b, "greaterThanOrEqual": lambda: a >= b, "lessThanOrEqual": lambda: a <= b,
             "between": lambda: min(b, b2) <= a <= max(b, b2),
             "notBetween": lambda: not min(b, b2) <= a <= max(b, b2)}
    try:
        return tests[op]()
    except (TypeError, KeyError):
        return False


def _percentile(vals, p):
    s = sorted(vals)
    k = (len(s) - 1) * p
    lo = math.floor(k)
    return s[lo] + (s[min(lo + 1, len(s) - 1)] - s[lo]) * (k - lo)


def _scale_color(colors, stops, v):
    rgb = lambda c: [int(c[i:i + 2], 16) for i in (1, 3, 5)]
    if v <= stops[0]:
        return colors[0]
    for i in range(1, len(stops)):
        if v <= stops[i]:
            span = stops[i] - stops[i - 1] or 1
            t = (v - stops[i - 1]) / span
            a, b = rgb(colors[i - 1]), rgb(colors[i])
            return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))
    return colors[-1]


def _apply_cf(ws, cells, styles, notes):
    rules = []
    for cf in ws.findall("conditionalFormatting"):
        ranges = [r for r in (parse_range(x) for x in (cf.get("sqref") or "").split()) if r]
        for rule in cf.findall("cfRule"):
            rules.append((int(rule.get("priority", "0")), ranges, rule))
    unsupported = set()
    for _, ranges, rule in sorted(rules, key=lambda x: x[0]):
        typ, op = rule.get("type"), rule.get("operator")
        dxf = styles.dxfs[int(rule.get("dxfId"))] if rule.get("dxfId") and int(rule.get("dxfId")) < len(styles.dxfs) else {}
        targets = [(rc, c) for rc, c in cells.items() if in_ranges(ranges, *rc)]
        if typ == "colorScale":
            cs = rule.find("colorScale")
            colors = [styles.color(c) or "#FFFFFF" for c in cs.findall("color")]
            vals = [c["num"] for _, c in targets if "num" in c]
            if not vals or len(colors) < 2:
                continue
            stops = []
            for cfvo in cs.findall("cfvo"):
                t, v = cfvo.get("type"), cfvo.get("val")
                try:
                    stops.append({"min": min(vals), "max": max(vals), "num": float(v or 0),
                                  "percent": min(vals) + (max(vals) - min(vals)) * float(v or 0) / 100,
                                  "percentile": _percentile(vals, float(v or 0) / 100)}[t])
                except (KeyError, ValueError):
                    stops.append(min(vals) if not stops else max(vals))
            for _, c in targets:
                if "num" in c:
                    c.setdefault("cf", {}).setdefault("bg", _scale_color(colors, stops, c["num"]))
            continue
        if typ not in ("cellIs", "containsText", "notContainsText", "beginsWith", "endsWith",
                       "containsBlanks", "notContainsBlanks"):
            unsupported.add({"expression": "formüllü kural", "dataBar": "veri çubuğu", "iconSet": "simge kümesi",
                             "top10": "ilk/son N", "aboveAverage": "ortalamanın üstü/altı",
                             "duplicateValues": "yinelenen değerler", "uniqueValues": "benzersiz değerler"}.get(typ, typ))
            continue
        try:
            fs = [_operand(f.text, cells) for f in rule.findall("formula")] if typ == "cellIs" else []
        except ValueError:
            unsupported.add("hücreye/formüle bağlı kural")
            continue
        text = (rule.get("text") or "").lower()
        for _, c in targets:
            v = c.get("num", c.get("str"))
            if typ == "cellIs":
                hit = v is not None and fs and _compare(v, op, fs[0], fs[1] if len(fs) > 1 else None)
            elif typ in ("containsBlanks", "notContainsBlanks"):
                hit = (v in (None, "")) == (typ == "containsBlanks")
            else:
                s = str(c.get("str", c.get("v", ""))).lower()
                hit = {"containsText": text in s, "notContainsText": text not in s,
                       "beginsWith": s.startswith(text), "endsWith": s.endswith(text)}[typ]
            if hit and not c.get("cfstop"):
                tgt = c.setdefault("cf", {})
                for k2, v2 in dxf.items():
                    tgt.setdefault(k2, v2)
                if rule.get("stopIfTrue") in ("1", "true"):
                    c["cfstop"] = True
    if unsupported:
        notes.append("Gösterilmeyen koşullu biçimler: " + ", ".join(sorted(unsupported)) + ".")


# ---- workbook -------------------------------------------------------------------
def load(path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return _load_csv(path)
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise ViewError("Dosya bir Excel (OOXML) paketi değil.")
    with z:
        names = set(z.namelist())
        root_rels = _rels(z, names, "")
        wb_part = next((r["part"] for r in root_rels if r["type"] == "officeDocument"), "xl/workbook.xml")
        if wb_part not in names:
            raise ViewError("Çalışma kitabı bulunamadı.")
        wb = _xml(z, wb_part)
        rels = {r["id"]: r for r in _rels(z, names, wb_part)}
        by_type = {r["type"]: r["part"] for r in rels.values()}
        theme = _theme(z, by_type.get("theme") if by_type.get("theme") in names else None)
        styles = Styles(_xml(z, by_type["styles"]) if by_type.get("styles") in names else None, theme)
        sst = []
        if by_type.get("sharedStrings") in names:
            sst = [_text(si) for si in _xml(z, by_type["sharedStrings"]).findall("si")]
        pr = wb.find("workbookPr")
        date1904 = pr is not None and pr.get("date1904") in ("1", "true")
        view = wb.find("bookViews/workbookView")
        active = int(view.get("activeTab", "0")) if view is not None else 0
        sheets = []
        for sh in wb.findall("sheets/sheet"):
            rel = rels.get(sh.get("id"))
            info = {"name": sh.get("name"), "state": sh.get("state", "visible")}
            if not rel or rel["part"] not in names or rel["type"] not in ("worksheet",):
                info.update(kind="chart" if rel and rel["type"] == "chartsheet" else "other", rows=0, cols=0, cells=[])
                sheets.append(info)
                continue
            info.update(_sheet(z, names, rel["part"], styles, sst, theme, date1904))
            sheets.append(info)
        if not sheets:
            raise ViewError("Çalışma kitabında sayfa yok.")
        if not (0 <= active < len(sheets)) or sheets[active]["state"] != "visible":
            active = next((i for i, s in enumerate(sheets) if s["state"] == "visible"), 0)
        return {"font": styles.default_font, "styles": styles.css, "sheets": sheets, "active": active}


def _sheet(z, names, part, styles, sst, theme, date1904):
    ws = _xml(z, part)
    notes, cells, rows, cols = [], {}, {}, []
    fmt_px = lambda w: max(0, round(float(w) * 7))
    fp = ws.find("sheetFormatPr")
    default_h = float(fp.get("defaultRowHeight", "15")) if fp is not None else 15.0
    if fp is not None and fp.get("defaultColWidth"):
        default_w = fmt_px(fp.get("defaultColWidth"))
    else:
        base = int(fp.get("baseColWidth", "8")) if fp is not None else 8
        default_w = math.ceil((base * 7 + 5) / 8) * 8
    for col in ws.findall("cols/col"):
        cols.append({"min": int(col.get("min")), "max": min(int(col.get("max")), MAX_COLS),
                     "w": fmt_px(col.get("width")) if col.get("width") else default_w,
                     "hidden": col.get("hidden") in ("1", "true"), "s": int(col.get("style", "0"))})
    view = ws.find("sheetViews/sheetView")
    grid = view is None or view.get("showGridLines") not in ("0", "false")
    freeze = None
    pane = view.find("pane") if view is not None else None
    if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
        freeze = {"rows": int(float(pane.get("ySplit", "0"))), "cols": int(float(pane.get("xSplit", "0")))}
    tab = ws.find("sheetPr/tabColor")
    shared, max_r, max_c, total, r_prev, cut_cols = {}, 0, 0, 0, 0, False
    for row in ws.findall("sheetData/row"):
        r = int(row.get("r")) if row.get("r") else r_prev + 1
        r_prev = r
        if r > MAX_ROWS:
            notes.append(f"Yalnızca ilk {MAX_ROWS} satır gösteriliyor.")
            break
        rinfo = {}
        if row.get("ht"):
            rinfo["h"] = float(row.get("ht"))
        if row.get("hidden") in ("1", "true"):
            rinfo["hidden"] = 1
        if row.get("customFormat") in ("1", "true") and row.get("s"):
            rinfo["s"] = int(row.get("s"))
        if rinfo:
            rows[r] = rinfo
            max_r = max(max_r, r) if "s" in rinfo else max_r
        c_prev = 0
        for c in row.findall("c"):
            ref = parse_ref(c.get("r")) if c.get("r") else (r, c_prev + 1)
            c_prev = ref[1]
            if ref[1] > MAX_COLS:
                cut_cols = True
                continue
            total += 1
            if total > MAX_CELLS:
                break
            cell = _cell(c, ref, styles, sst, shared, date1904)
            if cell:
                cells[ref] = cell
                max_r, max_c = max(max_r, ref[0]), max(max_c, ref[1])
        if total > MAX_CELLS:
            notes.append("Sayfa çok büyük; yalnızca bir kısmı gösteriliyor.")
            break
    if cut_cols:
        notes.append(f"Yalnızca ilk {MAX_COLS} sütun gösteriliyor.")
    merges = [m for m in (parse_range(x.get("ref")) for x in ws.findall("mergeCells/mergeCell")) if m]
    for r1, c1, r2, c2 in merges:
        max_r, max_c = max(max_r, min(r2, MAX_ROWS)), max(max_c, min(c2, MAX_COLS))

    rel_list = _rels(z, names, part)
    for rel in rel_list:
        if rel["type"] == "comments" and rel["part"] in names:
            for cm in _xml(z, rel["part"]).findall("commentList/comment"):
                ref = parse_ref(cm.get("ref"))
                if ref and ref[0] <= MAX_ROWS and ref[1] <= MAX_COLS:
                    cells.setdefault(ref, {"v": "", "k": ""})["note"] = _text(cm.find("text")).strip()
                    max_r, max_c = max(max_r, ref[0]), max(max_c, ref[1])
    if any(rel["type"] == "drawing" for rel in rel_list):
        drawing = next(rel["part"] for rel in rel_list if rel["type"] == "drawing")
        kinds = {r2["type"] for r2 in _rels(z, names, drawing)} if drawing in names else set()
        what = " ve ".join(x for x, ok in (("grafik", "chart" in kinds), ("görsel", "image" in kinds)) if ok) or "çizim"
        notes.append(f"Bu sayfada {what} var; tablo görünümünde gösterilmez — PDF görünümüne geçin.")

    tables = []
    for rel in rel_list:
        if rel["type"] != "table" or rel["part"] not in names:
            continue
        t = _xml(z, rel["part"])
        rng = parse_range(t.get("ref"))
        info = t.find("tableStyleInfo")
        look = _table_look(info.get("name") if info is not None else None, theme)
        if rng and look:
            tables.append({"range": rng, "header": int(t.get("headerRowCount", "1")),
                           "totals": int(t.get("totalsRowCount", "0")),
                           "stripes": info.get("showRowStripes") in ("1", "true"),
                           "first": info.get("showFirstColumn") in ("1", "true"), "look": look})
            max_r, max_c = max(max_r, min(rng[2], MAX_ROWS)), max(max_c, min(rng[3], MAX_COLS))

    validations = []
    for dv in ws.findall("dataValidations/dataValidation"):
        ranges = [x for x in (parse_range(s) for s in (dv.get("sqref") or "").split()) if x]
        typ, f1, f2 = dv.get("type", "any"), dv.findtext("formula1") or "", dv.findtext("formula2") or ""
        if typ == "list":
            text = "Liste: " + (f1.strip('"').replace(",", ", ") if f1.startswith('"') else f1)
        else:
            op = {"between": f"{f1} – {f2}", "notBetween": f"{f1} – {f2} dışı", "equal": f"= {f1}",
                  "notEqual": f"≠ {f1}", "greaterThan": f"> {f1}", "lessThan": f"< {f1}",
                  "greaterThanOrEqual": f"≥ {f1}", "lessThanOrEqual": f"≤ {f1}"}.get(dv.get("operator", "between"), f1)
            text = {"whole": "Tam sayı", "decimal": "Ondalık", "date": "Tarih", "time": "Saat",
                    "textLength": "Metin uzunluğu", "custom": "Formül"}.get(typ, typ) + f" {op}".rstrip()
        if dv.get("prompt"):
            text += f" · {dv.get('prompt')}"
        validations.append({"ranges": ranges, "text": text})

    _apply_cf(ws, cells, styles, notes)
    out_cells = []
    for (r, c), cell in sorted(cells.items()):
        cell.pop("num", None)
        cell.pop("str", None)
        cell.pop("cfstop", None)
        out_cells.append({"r": r, "c": c, **cell})
    return {"kind": "sheet", "rows": max_r, "cols": max_c, "cells": out_cells,
            "rowInfo": {str(k): v for k, v in rows.items()}, "colInfo": cols, "defaultH": default_h,
            "defaultW": default_w, "grid": grid, "freeze": freeze, "merges": merges, "tables": tables,
            "validations": validations, "tab": styles.color(tab) if tab is not None else None,
            "notes": list(dict.fromkeys(notes))}


def _cell(c, ref, styles, sst, shared, date1904):
    s = int(c.get("s", "0"))
    s = s if s < len(styles.css) else 0
    t, v = c.get("t", "n"), c.findtext("v")
    cell = {"s": s} if s else {}
    f = c.find("f")
    if f is not None:
        text = f.text
        if f.get("t") == "shared" and f.get("si") is not None:
            if text:
                shared[f.get("si")] = (text, ref)
            elif f.get("si") in shared:
                base, (r0, c0) = shared[f.get("si")]
                text = shift_formula(base, ref[0] - r0, ref[1] - c0)
        if text:
            cell["f"] = ("{=" + text + "}") if f.get("t") == "array" else "=" + text
    code = styles.fmt[s]
    if t == "inlineStr":
        cell.update(v=format_text(_text(c.find("is")), code), k="s", str=_text(c.find("is")))
    elif v is None:
        if not cell:
            return None
        cell.update(v="", k="")
    elif t == "s":
        txt = sst[int(v)] if v.isdigit() and int(v) < len(sst) else ""
        cell.update(v=format_text(txt, code), k="s", str=txt)
    elif t in ("str", "d"):
        cell.update(v=format_text(v, code), k="s", str=v)
    elif t == "b":
        cell.update(v="DOĞRU" if v in ("1", "true") else "YANLIŞ", k="b", num=1.0 if v in ("1", "true") else 0.0)
    elif t == "e":
        cell.update(v=v, k="e", str=v)
    else:
        try:
            x = float(v)
        except ValueError:
            cell.update(v=v, k="s", str=v)
            return cell
        text, color, is_date = format_number(x, code, date1904)
        cell.update(v=text, k="d" if is_date else "n", num=x)
        if color:
            cell["fc"] = color
        raw = text if is_date else general(x) if abs(x) < 1e15 else text
        if raw != text:
            cell["raw"] = raw
    return cell


def _load_csv(path):
    data = path.read_bytes()[:20 * 1024 * 1024]
    text = data.decode("utf-8-sig", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    cells, max_c, notes = [], 0, []
    num = re.compile(r"^-?\d{1,3}([.,]\d{3})*([.,]\d+)?$|^-?\d+([.,]\d+)?$")
    for r, row in enumerate(csv.reader(text.splitlines(), dialect), 1):
        if r > MAX_ROWS:
            notes.append(f"Yalnızca ilk {MAX_ROWS} satır gösteriliyor.")
            break
        for c, val in enumerate(row[:MAX_COLS], 1):
            if val != "":
                cells.append({"r": r, "c": c, "v": val, "k": "n" if num.match(val.strip()) else "s"})
                max_c = max(max_c, c)
    rows = cells[-1]["r"] if cells else 0
    sheet = {"name": path.stem, "state": "visible", "kind": "sheet", "rows": rows, "cols": max_c, "cells": cells,
             "rowInfo": {}, "colInfo": [], "defaultH": 15.0, "defaultW": 96, "grid": True, "freeze": None,
             "merges": [], "tables": [], "validations": [], "tab": None, "notes": notes}
    return {"font": {"fn": "Calibri", "sz": 11}, "styles": [{"fn": "Calibri", "sz": 11}], "sheets": [sheet], "active": 0}
