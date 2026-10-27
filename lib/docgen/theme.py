"""Theme colors/fonts for DOCX, PPTX and XLSX (DrawingML theme parts).

Changing a theme slot recolors everything bound to it (headings, accents,
charts) while explicit colors set by the author stay as they are.
"""
from lxml import etree

from . import NotFound

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
         "accent6", "hlink", "folHlink")


def a(tag):
    return f"{{{A_NS}}}{tag}"


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
    part (or only ``parts``)."""
    if slot not in SLOTS:
        raise NotFound(f"unknown theme slot {slot!r}; use one of {SLOTS}")
    hit = 0
    for part in parts or theme_parts(pkg):
        cs = pkg.xml(part).find(".//" + a("clrScheme"))
        el = cs.find(a(slot)) if cs is not None else None
        if el is None:
            continue
        for ch in list(el):
            el.remove(ch)
        etree.SubElement(el, a("srgbClr"), val=color.lstrip("#").upper())
        hit += 1
    if not hit:
        raise NotFound("no theme part with a color scheme")
    return hit


def set_fonts(pkg, major=None, minor=None, parts=None):
    """Theme heading (major) / body (minor) latin fonts."""
    for part in parts or theme_parts(pkg):
        fs = pkg.xml(part).find(".//" + a("fontScheme"))
        if fs is None:
            continue
        for tag, name in (("majorFont", major), ("minorFont", minor)):
            if name:
                fs.find(a(tag)).find(a("latin")).set("typeface", name)
