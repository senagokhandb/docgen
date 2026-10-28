"""PDF editing with pypdfium2 (pages + text) and pdfplumber (reading).

Text replacement keeps the rest of the page untouched:
  strategy "inplace": the existing text object is rewritten with its own
      embedded font — used when every character of the new text already
      appears in that font elsewhere in the document and the line still fits.
  strategy "redraw": the affected lines are removed and drawn again with the
      full font file of the same family (wrapping inside the paragraph box).
      Refused when that family is not installed or the text does not fit.
Every edit is verified by re-reading the result before it is written.
"""
import collections
import ctypes
import glob
import os
import re
import shutil
import subprocess
import tempfile

import pdfplumber
import pypdfium2 as pdfium
import pypdfium2.raw as raw

from . import Ambiguous, DocgenError, NotFound, NotSupported

BAND = 0.08  # top/bottom share of the page treated as header/footer band
FONT_DIRS = ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
             "/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts")]


def _norm(t):
    return re.sub(r"\s+", " ", t or "").strip()


def base_font(name):
    return re.sub(r"^[A-Z]{6}\+", "", name or "")


# ---- reading model ------------------------------------------------------
class Line:
    def __init__(self, page_no, d, height):
        self.page = page_no
        self.text = d["text"]
        self.x0, self.x1, self.top, self.bottom = d["x0"], d["x1"], d["top"], d["bottom"]
        chars = [c for c in d["chars"] if c["text"].strip()]
        self.chars = d["chars"]
        sig = collections.Counter((c["fontname"], round(c["size"], 1), _rgb(c.get("non_stroking_color")))
                                  for c in chars)
        (self.font, self.size, self.color), _ = sig.most_common(1)[0] if sig else ((None, 0, (0, 0, 0)), 0)
        self.mixed = len(sig) > 1
        self.band = "header" if self.bottom < height * BAND else ("footer" if self.top > height * (1 - BAND) else None)

    def __repr__(self):
        return f"<Line p{self.page} y={self.top:.1f} {self.text[:50]!r}{' ' + self.band if self.band else ''}>"


def _rgb(c):
    if c is None:
        return (0, 0, 0)
    c = tuple(c) if isinstance(c, (list, tuple)) else (c,)
    if len(c) == 1:
        return (round(c[0] * 255),) * 3
    if len(c) == 3:
        return tuple(round(x * 255) for x in c)
    if len(c) == 4:  # CMYK
        k = c[3]
        return tuple(round(255 * (1 - x) * (1 - k)) for x in c[:3])
    return (0, 0, 0)


def lines(path):
    """[[Line, ...] per page]"""
    out = []
    with pdfplumber.open(path) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            out.append([Line(i, d, pg.height) for d in pg.extract_text_lines(return_chars=True, strip=True)])
    return out


def page_texts(path, body_only=False):
    return ["\n".join(l.text for l in ls if not (body_only and l.band)) for ls in lines(path)]


class Match:
    def __init__(self, page, first, last, start, end, ls):
        self.page, self.first, self.last, self.start, self.end, self._lines = page, first, last, start, end, ls

    def __repr__(self):
        t = " ".join(l.text for l in self._lines[self.first:self.last + 1])
        return f"page {self.page} line {self.first + 1}: …{t[max(0, self.start - 25):self.end + 25]}…"


def find(path, text, *, page=None, band=None, near=None):
    """Occurrences of ``text`` (whitespace-insensitive, may span lines).
    band: None = body only, 'header'/'footer' = that band, 'any' = all.
    near: keep matches whose line(s) also contain this text (table row key)."""
    ms = _find(path, text, page=page, band=band)
    if near:
        ms = [m for m in ms if _norm(near) in " ".join(_norm(l.text) for l in m._lines[m.first:m.last + 1])]
    return ms


def _find(path, text, *, page=None, band=None):
    want = _norm(text)
    out = []
    for pno, ls in enumerate(lines(path), 1):
        if page and pno != page:
            continue
        sel = [l for l in ls if (band == "any") or (l.band == band)]
        joined, spans = "", []
        for l in sel:
            if joined:
                joined += " "
            spans.append((len(joined), len(joined) + len(_norm(l.text))))
            joined += _norm(l.text)
        for m in re.finditer(re.escape(want), joined):
            first = next(i for i, (a, b) in enumerate(spans) if a <= m.start() < b + 1)
            last = next(i for i, (a, b) in enumerate(spans) if a < m.end() <= b + 1)
            out.append(Match(pno, first, last, m.start() - spans[first][0], m.end() - spans[last][0], sel))
    return out


# ---- fonts --------------------------------------------------------------
def _style_of(name):
    n = name.lower()
    bold = bool(re.search(r"bold|black|heavy|semibold|demi", n))
    italic = bool(re.search(r"italic|oblique", n))
    return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic",
            (True, True): "BoldItalic"}[(bold, italic)]


# metric-compatible substitutes (same glyph widths) for fonts that PDFs often
# reference without embedding, or that cannot be redistributed
FONT_SUBSTITUTES = {"helvetica": "LiberationSans", "arial": "LiberationSans", "arialmt": "LiberationSans",
                    "times": "LiberationSerif", "timesroman": "LiberationSerif", "timesnewroman": "LiberationSerif",
                    "timesnewromanpsmt": "LiberationSerif", "courier": "LiberationMono", "couriernew": "LiberationMono",
                    "calibri": "Carlito", "cambria": "Caladea"}


def font_substitute(pdf_font_name):
    family = re.split(r"[-,]", base_font(pdf_font_name))[0].replace(" ", "").lower()
    return FONT_SUBSTITUTES.get(family)


def find_font_file(pdf_font_name):
    """Full font file for a PDF font name (e.g. 'BAAAAA+Caladea-Bold'): the
    same family, or its metric-compatible substitute (Helvetica -> Liberation
    Sans, Calibri -> Carlito, ...). None when neither is installed."""
    path = _font_file(pdf_font_name)
    sub = font_substitute(pdf_font_name)
    if path is None and sub:
        style = _style_of(base_font(pdf_font_name))
        path = _font_file(f"{sub}-{style}")
    return path


def _font_file(pdf_font_name):
    base = base_font(pdf_font_name)
    family = re.split(r"[-,]", base)[0].replace(" ", "")
    style = _style_of(base)
    key = lambda s: re.sub(r"[^a-z]", "", s.lower())
    cands = []
    for d in FONT_DIRS + os.environ.get("DOCGEN_FONT_DIRS", "").split(":"):
        if d and os.path.isdir(d):
            cands += glob.glob(os.path.join(d, "**", "*.[ot]tf"), recursive=True)
    wanted = key(family + style)
    for f in cands:
        stem = key(os.path.splitext(os.path.basename(f))[0])
        if stem == wanted or (style == "Regular" and stem == key(family)):
            return f
    if shutil.which("fc-match"):
        st = {"Regular": "Regular", "Bold": "Bold", "Italic": "Italic", "BoldItalic": "Bold Italic"}[style]
        res = subprocess.run(["fc-match", f"{family}:style={st}", "-f", "%{family}|%{file}"],
                             capture_output=True, text=True).stdout
        fam, _, path = res.partition("|")
        if key(fam.split(",")[0]) == key(family) and os.path.exists(path):
            return path
    return None


def _width(font_file, size):
    from PIL import ImageFont
    f = ImageFont.truetype(font_file, 1000)
    return lambda s: f.getlength(s) * size / 1000.0


# ---- pdfium helpers -----------------------------------------------------
def _wide(s):
    buf = (s + "\x00").encode("utf-16-le")
    return buf, ctypes.cast(ctypes.c_char_p(buf), ctypes.POINTER(raw.FPDF_WCHAR))


def _obj_text(obj, tp):
    n = raw.FPDFTextObj_GetText(obj.raw, tp.raw, None, 0)
    buf = (ctypes.c_ushort * max(n // 2, 1))()
    raw.FPDFTextObj_GetText(obj.raw, tp.raw, buf, n)
    return bytes(buf).decode("utf-16-le", "ignore").rstrip("\x00")


def _text_objects(page):
    tp = page.get_textpage()
    out = []
    for o in page.get_objects(max_depth=1):
        if o.type == raw.FPDF_PAGEOBJ_TEXT:
            l, b, r, t = (ctypes.c_float() for _ in range(4))
            raw.FPDFPageObj_GetBounds(o.raw, l, b, r, t)
            l, b, r, t = l.value, b.value, r.value, t.value
            m = raw.FS_MATRIX()
            raw.FPDFPageObj_GetMatrix(o.raw, m)
            out.append({"obj": o, "text": _obj_text(o, tp), "box": (l, b, r, t), "origin": (m.e, m.f)})
    return out


def _objs_for_line(objs, line, height):
    top, bottom = height - line.top, height - line.bottom
    hits = []
    for o in objs:
        l, b, r, t = o["box"]
        cy = (b + t) / 2
        if bottom - 1 <= cy <= top + 1 and r > line.x0 - 1 and l < line.x1 + 1:
            hits.append(o)
    return sorted(hits, key=lambda o: o["box"][0])


def _chars_by_font(all_lines):
    seen = collections.defaultdict(set)
    widths = collections.defaultdict(list)
    for ls in all_lines:
        for l in ls:
            for c in l.chars:
                k = (c["fontname"], round(c["size"], 1))
                seen[c["fontname"]].add(c["text"])
                widths[(k, c["text"])].append(c["x1"] - c["x0"])
    return seen, {k: sum(v) / len(v) for k, v in widths.items()}


# ---- text replacement ---------------------------------------------------
def _segments(line_objs, size):
    """Split a line's text objects into segments (table cells / runs of text)
    wherever the horizontal gap is larger than a word space."""
    segs = []
    for o in line_objs:
        if segs and o["box"][0] - segs[-1][-1]["box"][2] <= size * 1.2:
            segs[-1].append(o)
        else:
            segs.append([o])
    return segs


def _seg_text(seg):
    return "".join(o["text"] for o in seg)


def _column_align(objs, page_lines, seg, height, size):
    """left/right/center alignment of a table column, judged from the
    segments of other lines that overlap it horizontally."""
    x0, x1 = seg[0]["box"][0], seg[-1]["box"][2]
    votes = collections.Counter()
    for l in page_lines:
        for sg in _segments(_objs_for_line(objs, l, height), size):
            a, b = sg[0]["box"][0], sg[-1]["box"][2]
            if sg is seg or b < x0 - 1 or a > x1 + 1 or (a, b) == (x0, x1):
                continue
            if abs(a - x0) < 1:
                votes["left"] += 1
            elif abs(b - x1) < 1:
                votes["right"] += 1
            elif abs((a + b) / 2 - (x0 + x1) / 2) < 1:
                votes["center"] += 1
    return votes.most_common(1)[0][0] if votes else "left"


def _key(line):
    return (round(line.top, 1), line.text)


def replace_text(src, dst, old, new, *, page=None, occurrence=None, band=None, near=None, strategy="auto"):
    """Replace one occurrence of ``old`` with ``new``. Narrow the target with
    page=, band='header'|'footer', near='row key' or occurrence=n (1-based);
    several matches otherwise raise Ambiguous. Returns a report dict."""
    ms = find(src, old, page=page, band=band, near=near)
    if not ms:
        raise NotFound(f"{old!r} not found" + (f" on page {page}" if page else ""))
    if occurrence:
        ms = [ms[occurrence - 1]]
    if len(ms) > 1:
        raise Ambiguous(f"{old!r}", ms)
    m = ms[0]
    all_lines = lines(src)
    page_lines = all_lines[m.page - 1]
    sel = m._lines
    block = sel[m.first:m.last + 1]
    seen, widths = _chars_by_font(all_lines)
    pdf = pdfium.PdfDocument(src)
    pg = pdf[m.page - 1]
    height, pwidth = pg.get_height(), pg.get_width()
    objs = _text_objects(pg)
    f0 = block[0]
    report = {"page": m.page, "old": old, "new": new, "lines_before": [l.text for l in block]}

    def est(text):
        total = 0.0
        for ch in text:
            w = widths.get(((f0.font, f0.size), ch))
            if w is None:
                return None
            total += w
        return total

    def untouched(changed):
        keys = {_key(l) for l in changed}
        return [l.text for l in page_lines if _key(l) not in keys]

    line_objs = [_objs_for_line(objs, l, height) for l in block]
    if any(not lo for lo in line_objs):
        raise NotSupported("text is not drawn by top-level text objects (form XObject/annotation)")
    if re.sub(r"\s", "", "".join(o["text"] for lo in line_objs for o in lo)) != \
            re.sub(r"\s", "", "".join(l.text for l in block)):
        raise NotSupported("text objects do not line up with the visible lines; refusing to edit")
    chars_ok = set(new) - {" "} <= seen[f0.font]

    # 1) in place inside one segment (whole line, table cell, header piece)
    if m.first == m.last and strategy in ("auto", "inplace") and chars_ok and not f0.mixed:
        segs = [sg for sg in _segments(line_objs[0], f0.size) if _norm(old) in _norm(_seg_text(sg))]
        if len(segs) == 1:
            seg = segs[0]
            old_t = _norm(_seg_text(seg))
            new_t = old_t.replace(_norm(old), new, 1)
            w_old, w_new = seg[-1]["box"][2] - seg[0]["box"][0], est(new_t)
            col_right = max((l.x1 for l in page_lines if abs(l.x0 - f0.x0) < 2), default=pwidth - f0.x0)
            centered = _centered(f0, pwidth)
            room = (col_right - seg[0]["box"][0] + 0.5) if len(_segments(line_objs[0], f0.size)) == 1 else w_old * 1.15
            if w_new is not None and (w_new <= room or centered):
                _buf, ptr = _wide(new_t)
                if not raw.FPDFText_SetText(seg[0]["obj"].raw, ptr):
                    raise DocgenError("pdfium refused the text")
                for o in seg[1:]:
                    pg.remove_obj(o["obj"])
                align = _column_align(objs, page_lines, seg, height, f0.size)
                if centered and len(_segments(line_objs[0], f0.size)) == 1:
                    align = "center"
                dx = {"right": w_old - w_new, "center": (w_old - w_new) / 2}.get(align, 0)
                if dx:
                    raw.FPDFPageObj_Transform(seg[0]["obj"].raw, 1, 0, 0, 1, dx, 0)
                report.update(strategy="inplace", segment_before=old_t, segment_after=new_t)
                return _finish(pdf, pg, src, dst, m, new_t, report, untouched(block))
        if strategy == "inplace":
            raise NotSupported("in-place edit not possible (text spans segments or does not fit)")
    elif strategy == "inplace":
        raise NotSupported("characters of the new text are not in the embedded font subset"
                           if not chars_ok else "line mixes formatting; in-place edit refused")

    # 1b) one table cell / segment of a multi-column line: redraw only that
    #     segment with the full font file, keeping the column alignment
    all_segs = _segments(line_objs[0], f0.size) if m.first == m.last else []
    cell = [sg for sg in all_segs if _norm(old) in _norm(_seg_text(sg))]
    if len(all_segs) > 1 and len(cell) == 1 and strategy in ("auto", "redraw"):
        seg = cell[0]
        x0, x1 = seg[0]["box"][0], seg[-1]["box"][2]
        chars = [c for c in block[0].chars if c["text"].strip() and c["x0"] >= x0 - 0.5 and c["x1"] <= x1 + 0.5]
        sig = collections.Counter((c["fontname"], round(c["size"], 1), _rgb(c.get("non_stroking_color"))) for c in chars)
        if len(sig) != 1:
            raise NotSupported("the cell mixes formatting; redrawing it would lose that formatting")
        (font, size, color), _ = sig.most_common(1)[0]
        font_file = find_font_file(font)
        if not font_file:
            raise NotSupported(f"font family of {base_font(font)!r} is not installed; cannot redraw the cell")
        new_t = _norm(_seg_text(seg)).replace(_norm(old), new, 1)
        w_new = _width(font_file, size)(new_t)
        i = all_segs.index(seg)
        left = all_segs[i - 1][-1]["box"][2] + size * 0.5 if i > 0 else 0
        right = all_segs[i + 1][0]["box"][0] - size * 0.5 if i + 1 < len(all_segs) else pwidth - f0.x0
        align = _column_align(objs, page_lines, seg, height, size)
        bx, by = seg[0]["origin"]
        nx = {"right": bx + (x1 - x0) - w_new, "center": bx + ((x1 - x0) - w_new) / 2}.get(align, bx)
        if nx < left or nx + w_new > right:
            raise NotSupported("the new cell text does not fit in its column")
        for o in seg:
            pg.remove_obj(o["obj"])
        keep = _draw(pdf, pg, font_file, size, color, [(nx, by, new_t)])
        report.update(strategy="redraw-cell", font_file=font_file, segment_after=new_t, align=align,
                      font_substitute=font_substitute(font) if base_font(font).split("-")[0].lower() not in
                      font_file.lower() else None)
        result = _finish(pdf, pg, src, dst, m, new_t, report, untouched(block))
        del keep
        return result

    # 2) redraw the paragraph with the full font file of the same family
    if any(len(_segments(lo, f0.size)) > 1 for lo in line_objs):
        raise NotSupported("the line is a table row / has separate columns; target one cell "
                           "(old=<cell text>, near=<row key>) instead of the whole row")
    if any(l.mixed for l in block):
        raise NotSupported("the text mixes fonts/sizes/colors; redrawing it would lose that formatting")
    font_file = find_font_file(f0.font)
    if not font_file:
        raise NotSupported(f"font family of {base_font(f0.font)!r} is not installed; cannot redraw the text "
                           f"without changing its look")
    width = _width(font_file, f0.size)
    para = list(block)
    pitch = (block[1].top - block[0].top) if len(block) > 1 else None
    if not f0.band:
        idx = sel.index(block[-1])
        for l in sel[idx + 1:]:
            gap = l.top - para[-1].top
            if abs(l.x0 - f0.x0) < 2 and abs(l.size - f0.size) < 0.2 and l.font == f0.font \
                    and gap < f0.size * 1.6 and (pitch is None or abs(gap - pitch) < 1.5):
                pitch = pitch or gap
                para.append(l)
            else:
                break
    pitch = pitch or f0.size * 1.2
    col_right = max((l.x1 for l in page_lines if abs(l.x0 - f0.x0) < 2 and not l.band), default=pwidth - f0.x0)
    if len(para) == 1:  # a single line (heading, caption) may grow up to the right margin
        col_right = max(col_right, pwidth - f0.x0)
    text = " ".join([joined_new(block, m, new)] + [_norm(l.text) for l in para[len(block):]])
    wrapped = _wrap(text, width, col_right - f0.x0 + 0.5)
    if f0.band and len(wrapped) > 1:
        raise NotSupported("new header/footer text does not fit on one line")
    nxt = next((l for l in page_lines if l.top > para[-1].top + 0.5 and _key(l) not in {_key(p) for p in para}), None)
    limit = (nxt.top - f0.top) if nxt else (height * (1 - BAND) - f0.top)
    if (len(wrapped) - 1) * pitch + f0.size > limit:
        raise NotSupported(f"new text needs {len(wrapped)} lines; only room for "
                           f"{int((limit - f0.size) // pitch) + 1} before the next block")
    para_objs = [_objs_for_line(objs, l, height) for l in para]
    base_x, base_y = para_objs[0][0]["origin"]
    for lo in para_objs:
        for o in lo:
            pg.remove_obj(o["obj"])
    keep = _draw(pdf, pg, font_file, f0.size, f0.color, [(base_x, base_y - i * pitch, t) for i, t in enumerate(wrapped)])
    report.update(strategy="redraw", font_file=font_file, lines_after=wrapped,
                  font_substitute=font_substitute(f0.font) if base_font(f0.font).split("-")[0].lower() not in
                  font_file.lower() else None)
    result = _finish(pdf, pg, src, dst, m, text, report, untouched(para))
    del keep
    return result


def _draw(pdf, pg, font_file, size, color, items):
    """Insert text objects [(x, baseline_y, text)] using a full TrueType font.
    Returns buffers that must stay alive until the document is saved."""
    with open(font_file, "rb") as fh:
        data = fh.read()
    fbuf = (ctypes.c_uint8 * len(data)).from_buffer_copy(data)
    font = raw.FPDFText_LoadFont(pdf.raw, fbuf, len(data), raw.FPDF_FONT_TRUETYPE, True)
    r, g, b = color
    keep = [fbuf]
    for x, y, text in items:
        t = raw.FPDFPageObj_CreateTextObj(pdf.raw, font, ctypes.c_float(size))
        _buf, ptr = _wide(text)
        keep.append(_buf)
        raw.FPDFText_SetText(t, ptr)
        raw.FPDFPageObj_SetFillColor(t, r, g, b, 255)
        raw.FPDFPageObj_Transform(t, 1, 0, 0, 1, x, y)
        raw.FPDFPage_InsertObject(pg.raw, t)
    return keep


def joined_new(block, m, new):
    joined = " ".join(_norm(l.text) for l in block)
    return joined[:m.start] + new + joined[m.end:]


def _centered(line, page_width):
    return abs((line.x0 + line.x1) / 2 - page_width / 2) < 3 and line.x0 > page_width * 0.2


def _wrap(text, width, limit):
    out, cur = [], ""
    for word in text.split(" "):
        cand = (cur + " " + word) if cur else word
        if width(cand) <= limit or not cur:
            cur = cand
        else:
            out.append(cur)
            cur = word
    if cur:
        out.append(cur)
    return out


def _finish(pdf, pg, src, dst, m, expected, report, untouched=()):
    pg.gen_content()
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    pdf.save(tmp)
    before, after = page_texts(src), page_texts(tmp)
    if len(before) != len(after):
        raise DocgenError("page count changed during a text edit")
    others = [i + 1 for i, (a, b) in enumerate(zip(before, after)) if a != b and i + 1 != m.page]
    if others:
        raise DocgenError(f"verification failed: pages {others} changed")
    page_after = after[m.page - 1].split("\n")
    if _norm(expected) not in _norm(" ".join(page_after)):
        raise DocgenError("verification failed: the new text cannot be read back from the page")
    lost = [t for t in untouched if t not in page_after]
    if lost:
        raise DocgenError(f"verification failed: other lines on page {m.page} changed: {lost[:3]}")
    shutil.move(tmp, dst)
    report["verified"] = True
    return report


def insert_paragraphs(src, dst, after, paragraphs, *, page=None, near=None):
    """Add paragraphs right below the line containing ``after``, in the body
    font, size and color of the page and inside the text column. No reflow:
    if they do not fit before the next block / the footer band the edit is
    refused; the user decides what happens to the text that does not fit."""
    ms = find(src, after, page=page, near=near)
    if not ms:
        raise NotFound(f"{after!r} not found")
    if len(ms) > 1:
        raise Ambiguous(f"{after!r}", ms)
    m = ms[0]
    page_lines = lines(src)[m.page - 1]
    anchor = m._lines[m.last]
    body = [l for l in page_lines if not l.band and not l.mixed]
    style = collections.Counter((l.font, l.size, l.color) for l in body).most_common(1)[0][0]
    font_name, size, color = style
    font_file = find_font_file(font_name)
    if not font_file:
        raise NotSupported(f"font family of {base_font(font_name)!r} is not installed")
    pdf = pdfium.PdfDocument(src)
    pg = pdf[m.page - 1]
    height, pwidth = pg.get_height(), pg.get_width()
    left = min(l.x0 for l in body)
    right = max(l.x1 for l in body if abs(l.x0 - left) < 2)
    if right - left < pwidth * 0.4:
        right = pwidth - left
    tops = sorted(l.top for l in body if abs(l.size - size) < 0.2)
    gaps = sorted(b - a for a, b in zip(tops, tops[1:]) if 0 < b - a < size * 1.6)
    pitch = gaps[len(gaps) // 2] if gaps else size * 1.2
    width = _width(font_file, size)
    anchor_objs = _objs_for_line(_text_objects(pg), anchor, height)
    y = (anchor_objs[0]["origin"][1] if anchor_objs else height - anchor.bottom) - pitch * 1.9
    items, para_last_y = [], []
    for i, ptext in enumerate(paragraphs):
        for line in _wrap(_norm(ptext), width, right - left):
            items.append((left, y, line))
            y -= pitch
        para_last_y.append(items[-1][1])
        y -= pitch * 0.9
    below = [l for l in page_lines if l.top > anchor.bottom + 0.5 and not l.band]
    floor = (height - below[0].top + 2) if below else height * BAND + size
    if items and items[-1][1] < floor:
        # same unit for both numbers: line pitches, paragraph gaps included
        need = round((items[0][1] - items[-1][1]) / pitch) + 1
        room = max(int((items[0][1] - floor) // pitch) + 1, 0) if items[0][1] > floor else 0
        fit = sum(1 for ly in para_last_y if ly >= floor)
        raise NotSupported(f"the {len(paragraphs)} paragraphs need about {need} lines of space (paragraph gaps "
                           f"included) but only about {room} are free on page {m.page} before the next "
                           f"block/footer — {fit} of {len(paragraphs)} would fit. No reflow: write no file, do "
                           f"not split, shrink or move text yourself, and do not add pages; ask the user what "
                           f"should happen to the text that does not fit (e.g. a new page after page {m.page}, a "
                           f"shorter text, or the source document). Where the request puts the text (\"at the "
                           f"end of section X\") is the insertion point, not that answer")
    keep = _draw(pdf, pg, font_file, size, color, items)
    pg.gen_content()
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    pdf.save(tmp)
    del keep
    before, after_t = page_texts(src), page_texts(tmp)
    others = [i + 1 for i, (a_, b_) in enumerate(zip(before, after_t)) if a_ != b_ and i + 1 != m.page]
    lost = [l.text for l in page_lines if l.text not in after_t[m.page - 1].split("\n")]
    flat = lambda s: re.sub(r"\s", "", s)
    missing = [p[:30] for p in paragraphs if flat(_norm(p))[:40] not in flat(after_t[m.page - 1])]
    if others or lost or missing or len(before) != len(after_t):
        raise DocgenError(f"verification failed: pages {others} changed, lost lines {lost[:2]}, missing {missing[:2]}")
    shutil.move(tmp, dst)
    return {"page": m.page, "after": anchor.text, "lines_added": len(items), "font_file": font_file,
            "font": base_font(font_name), "size": size, "verified": True}


def renumber(src, dst, total=None):
    """Rewrite 'n / N'-style page numbers in header/footer bands after pages
    were removed, inserted or moved (digits are replaced in place)."""
    pdf = pdfium.PdfDocument(src)
    n_pages = len(pdf)
    total = total or n_pages
    all_lines = lines(src)
    seen, _ = _chars_by_font(all_lines)
    changed = []
    for pno, ls in enumerate(all_lines, 1):
        pg = pdf[pno - 1]
        objs = None
        for l in ls:
            if not l.band:
                continue
            if not re.search(r"(?<![\d.])\d{1,4}\s*(?:/|of)\s*\d{1,4}(?![\d.])", l.text):
                continue
            new_nums = [str(pno), str(total)]
            objs = objs or _text_objects(pg)
            lo = _objs_for_line(objs, l, pg.get_height())
            queue = list(new_nums)
            for o in lo:
                if not re.search(r"\d", o["text"]):
                    continue
                new_t = re.sub(r"\d+", lambda _m: queue.pop(0) if queue else _m.group(0), o["text"])
                if set(new_t) - {" "} - seen[l.font]:
                    raise NotSupported(f"page {pno}: digits {new_t!r} not in the embedded font")
                if new_t != o["text"]:
                    _buf, ptr = _wide(new_t)
                    raw.FPDFText_SetText(o["obj"].raw, ptr)
            pg.gen_content()
            changed.append((pno, l.text))
    pdf.save(dst)
    return {"pages": n_pages, "renumbered": changed}


# ---- page operations ----------------------------------------------------
def _pages_arg(spec, n):
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += list(range(int(a), int(b) + 1))
        elif part:
            out.append(int(part))
    bad = [p for p in out if not 1 <= p <= n]
    if bad:
        raise NotFound(f"pages {bad} out of range 1..{n}")
    return out


def delete_pages(src, dst, pages):
    pdf = pdfium.PdfDocument(src)
    for p in sorted(set(_pages_arg(pages, len(pdf))), reverse=True):
        pdf.del_page(p - 1)
    pdf.save(dst)
    return len(pdf)


def keep_pages(src, dst, pages):
    pdf = pdfium.PdfDocument(src)
    keep = set(_pages_arg(pages, len(pdf)))
    return delete_pages(src, dst, ",".join(str(p) for p in range(1, len(pdf) + 1) if p not in keep))


def move_pages(src, dst, pages, before):
    """Move pages (e.g. '5' or '5-6') so they stand before page ``before``
    (original numbering; before = n+1 means 'to the end')."""
    pdf = pdfium.PdfDocument(src)
    n = len(pdf)
    sel = _pages_arg(pages, n)
    rest = [p for p in range(1, n + 1) if p not in sel]
    pos = len([p for p in rest if p < before])
    order = rest[:pos] + sel + rest[pos:]
    idx = (ctypes.c_int * len(sel))(*[p - 1 for p in sel])
    if not raw.FPDF_MovePages(pdf.raw, idx, len(sel), pos):
        raise DocgenError("pdfium could not move the pages")
    pdf.save(dst)
    return order


def insert_pdf(src, dst, other, at=None, pages=None):
    """Insert pages of ``other`` before page ``at`` (default: append)."""
    pdf = pdfium.PdfDocument(src)
    oth = pdfium.PdfDocument(other)
    sel = [p - 1 for p in _pages_arg(pages, len(oth))] if pages else None
    pdf.import_pages(oth, pages=sel, index=(at - 1) if at else None)
    pdf.save(dst)
    return len(pdf)


def rotate_pages(src, dst, pages, degrees):
    pdf = pdfium.PdfDocument(src)
    for p in _pages_arg(pages, len(pdf)):
        pg = pdf[p - 1]
        pg.set_rotation((pg.get_rotation() + degrees) % 360)
    pdf.save(dst)
