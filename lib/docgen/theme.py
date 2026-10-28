"""Theme colors/fonts for DOCX, PPTX and XLSX (DrawingML theme parts).

Changing a theme slot recolors everything bound to it (headings, accents,
charts) while explicit colors set by the author stay as they are.

DOCX keeps a cached hex value next to every theme reference
(``<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>``).
Word recomputes it from the theme; LibreOffice (and so Gotenberg, PDF export
and render_check) draws the cached value. ``set_color`` therefore refreshes
these caches too, or the document would still look unchanged outside Word.
"""
import colorsys
import math
import re

from lxml import etree

from . import NotFound

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
         "accent6", "hlink", "folHlink")

# WordprocessingML theme color name -> w:clrSchemeMapping attribute (settings.xml);
# text/background names go through the document's mapping (ECMA-376 17.15.1.20)
_W_MAPPED = {"text1": "t1", "background1": "bg1", "text2": "t2", "background2": "bg2"}
_W_DEFAULT_MAP = {"t1": "dark1", "bg1": "light1", "t2": "dark2", "bg2": "light2"}
# ST_ColorSchemeIndex -> DrawingML clrScheme slot (accent1..6 keep their name)
_W_SLOT = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
           "hyperlink": "hlink", "followedHyperlink": "folHlink"}
# (theme attribute, tint attribute, shade attribute, cache attribute; None = w:val
# on <w:color>, w:color on borders / <w:u> / <w:shd>)
_W_REFS = (("themeColor", "themeTint", "themeShade", None),
           ("themeFill", "themeFillTint", "themeFillShade", "fill"))


def a(tag):
    return f"{{{A_NS}}}{tag}"


def w(tag):
    return f"{{{W_NS}}}{tag}"


def theme_parts(pkg):
    return [n for n in pkg.names() if n.startswith(("word/theme/", "ppt/theme/", "xl/theme/")) and n.endswith(".xml")]


def get_colors(pkg):
    out = {}
    for part in theme_parts(pkg):
        cs = pkg.xml(part).find(".//" + a("clrScheme"))
        if cs is None:
            continue
        for slot in cs:
            c = slot[0] if len(slot) else None
            if c is not None:
                out.setdefault(part, {})[etree.QName(slot).localname] = c.get("val") or c.get("lastClr")
    return out


def set_color(pkg, slot, color, parts=None):
    """Set a theme color slot (accent1, dk2, ...) to 'RRGGBB' in every theme
    part (or only ``parts``). In a DOCX the cached hex values bound to that
    slot are recomputed as well (see the module docstring)."""
    if slot not in SLOTS:
        raise NotFound(f"unknown theme slot {slot!r}; use one of {SLOTS}")
    hit = []
    for part in parts or theme_parts(pkg):
        cs = pkg.xml(part).find(".//" + a("clrScheme"))
        el = cs.find(a(slot)) if cs is not None else None
        if el is None:
            continue
        for ch in list(el):
            el.remove(ch)
        etree.SubElement(el, a("srgbClr"), val=color.lstrip("#").upper())
        hit.append(part)
    if not hit:
        raise NotFound("no theme part with a color scheme")
    if _word_theme_part(pkg) in hit:
        refresh_word_caches(pkg, {slot})
    return len(hit)


def _word_theme_part(pkg):
    if not pkg.has("word/document.xml"):
        return None
    return next((r["part"] for r in pkg.rels("word/document.xml") if r.get("Type") == THEME_REL), None)


def _slot_rgb(cs, slot):
    el = cs.find(a(slot))
    c = el[0] if el is not None and len(el) else None
    if c is None:
        return None
    val = c.get("lastClr") if etree.QName(c).localname == "sysClr" else c.get("val")
    return val.upper() if val and re.fullmatch(r"[0-9A-Fa-f]{6}", val) else None


def derive(rgb, tint=None, shade=None):
    """Word's w:themeTint / w:themeShade ('00'..'FF') applied to 'RRGGBB': the
    HSL lightness is scaled (shade: L*s, tint: L*t + 1-t; s, t = hex/255) and
    rounded down; with both, Word applies the tint ([MS-OI29500] 2.1.72).
    Reproduces every cache Word wrote in the eval fixtures to within 1/255
    per channel."""
    if not (tint or shade):
        return rgb.upper()
    h, light, s = colorsys.rgb_to_hls(*(int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4)))
    x = int(tint or shade, 16) / 255
    light = light * x + 1 - x if tint else light * x
    return "".join(f"{min(255, math.floor(v * 255 + 1e-9)):02X}" for v in colorsys.hls_to_rgb(h, light, s))


def refresh_word_caches(pkg, slots=None):
    """Recompute the cached hex values of theme-bound colors in every
    WordprocessingML part from the document's current theme. Only references
    to ``slots`` (DrawingML names; None = all) are touched, and only caches
    that already exist. Returns the number of values changed."""
    theme = _word_theme_part(pkg)
    cs = pkg.xml(theme).find(".//" + a("clrScheme")) if theme and pkg.has(theme) else None
    if cs is None:
        return 0
    mapping = dict(_W_DEFAULT_MAP)
    if pkg.has("word/settings.xml"):
        m = pkg.xml("word/settings.xml").find(w("clrSchemeMapping"))
        for k, v in (m.attrib.items() if m is not None else ()):
            mapping[etree.QName(k).localname] = v

    def slot_of(name):
        name = mapping.get(_W_MAPPED.get(name, name), name)
        return _W_SLOT.get(name, name)

    changed = 0
    for part in pkg.names("word/*.xml"):  # fnmatch: also word/glossary/..., headers, footnotes
        data = pkg.read(part)
        if part.startswith("word/theme/") or not (b"themeColor" in data or b"themeFill" in data):
            continue
        for el in pkg.xml(part).iter(etree.Element):
            for ref, tint, shade, cache in _W_REFS:
                name = el.get(w(ref))
                if not name or name == "none":
                    continue
                slot = slot_of(name)
                base = _slot_rgb(cs, slot) if not slots or slot in slots else None
                cache = w(cache or ("val" if el.tag == w("color") else "color"))
                old = el.get(cache)
                if not base or not old or old == "auto":
                    continue
                new = derive(base, el.get(w(tint)), el.get(w(shade)))
                if new != old.upper():
                    el.set(cache, new)
                    changed += 1
    return changed


def set_fonts(pkg, major=None, minor=None, parts=None):
    """Theme heading (major) / body (minor) latin fonts."""
    for part in parts or theme_parts(pkg):
        fs = pkg.xml(part).find(".//" + a("fontScheme"))
        if fs is None:
            continue
        for tag, name in (("majorFont", major), ("minorFont", minor)):
            if name:
                fs.find(a(tag)).find(a("latin")).set("typeface", name)
