"""OOXML package I/O that keeps untouched parts byte-identical.

    from docgen.package import Package
    pkg = Package("/workspace/uploads/rapor.docx")
    root = pkg.xml("word/document.xml")      # lxml element, cached
    ...modify root...
    pkg.touch("word/document.xml")            # mark as modified
    pkg.save("/workspace/outputs/rapor.docx")

Only touched parts are re-serialized. Everything else is copied as-is, in
the original entry order, without ZIP directory entries.
"""
import copy
import fnmatch
import posixpath
import shutil
import zipfile

from lxml import etree

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
_BLANK_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True,
                                remove_blank_text=True)

# template / macro main parts -> plain document main part
_MAIN_SWITCH = {
    b"wordprocessingml.template.main+xml": b"wordprocessingml.document.main+xml",
    b"presentationml.template.main+xml": b"presentationml.presentation.main+xml",
    b"spreadsheetml.template.main+xml": b"spreadsheetml.sheet.main+xml",
}

KIND_BY_MAIN = {
    "word/document.xml": "docx",
    "ppt/presentation.xml": "pptx",
    "xl/workbook.xml": "xlsx",
}


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Package:
    def __init__(self, path):
        self.path = str(path)
        with zipfile.ZipFile(self.path) as z:
            infos = [i for i in z.infolist() if not i.is_dir()]
            self._data = {i.filename: z.read(i.filename) for i in infos}
            self._info = {i.filename: i for i in infos}
        self._order = list(self._data)
        self._orig = dict(self._data)
        self._xml = {}
        self._orig_c14n = {}
        self._touched = set()

    # ---- basic access -------------------------------------------------
    @property
    def kind(self):
        for main, kind in KIND_BY_MAIN.items():
            if main in self._data:
                return kind
        return None

    def names(self, pattern="*"):
        return [n for n in self._order if fnmatch.fnmatch(n, pattern)]

    def has(self, name):
        return name in self._data

    def read(self, name):
        """Current bytes of a part. A part loaded with xml() is re-serialized
        only if its tree really differs from the original."""
        if name in self._xml:
            root = self._xml[name]
            if name in self._touched or name not in self._orig_c14n:
                return serialize(root)
            if etree.tostring(root, method="c14n") != self._orig_c14n[name]:
                return serialize(root)
        return self._data[name]

    def xml(self, name):
        """Parsed root of an XML part (cached). Edits are detected on save."""
        if name not in self._xml:
            root = etree.fromstring(self._data[name], PARSER)
            self._xml[name] = root
            self._orig_c14n[name] = etree.tostring(root, method="c14n")
        return self._xml[name]

    def touch(self, name):
        """Optional: force re-serialization of a part."""
        if name not in self._xml:
            raise KeyError(f"{name} was not loaded with xml()")
        self._touched.add(name)

    def modified(self):
        """Parts whose bytes differ from the opened file (incl. new parts)."""
        return [n for n in self._order if self.read(n) != self._orig.get(n)]

    def removed(self):
        return [n for n in self._orig if n not in self._data]

    def write(self, name, data):
        """Add or replace a part with raw bytes (or an lxml element)."""
        if isinstance(data, etree._Element):
            self._xml[name] = data
            self._touched.add(name)
            data = serialize(data)
        else:
            self._xml.pop(name, None)
            self._orig_c14n.pop(name, None)
            self._touched.discard(name)
        if name not in self._data:
            self._order.append(name)
        self._data[name] = data

    def remove(self, name):
        self._data.pop(name, None)
        self._xml.pop(name, None)
        self._orig_c14n.pop(name, None)
        self._touched.discard(name)
        if name in self._order:
            self._order.remove(name)

    def save(self, dst):
        dst = str(dst)
        if not self.modified() and not self.removed() and list(self._orig) == self._order:
            if dst != self.path:
                shutil.copyfile(self.path, dst)  # nothing changed: byte-identical file
            return dst
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
            for name in self._order:
                data = self.read(name)
                old = self._info.get(name)
                if old is not None:
                    info = zipfile.ZipInfo(name, date_time=old.date_time)
                    info.compress_type = old.compress_type
                    info.external_attr = old.external_attr
                else:
                    info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                zo.writestr(info, data)
        return dst

    # ---- relationships ------------------------------------------------
    @staticmethod
    def rels_name(part):
        d, b = posixpath.split(part)
        return posixpath.join(d, "_rels", b + ".rels")

    @staticmethod
    def resolve(part, target):
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))

    def rels(self, part):
        """[{Id, Type, Target, TargetMode, part}] — part = resolved internal target."""
        name = self.rels_name(part)
        if not self.has(name):
            return []
        out = []
        for r in self.xml(name):
            d = dict(r.attrib)
            if d.get("TargetMode") != "External":
                d["part"] = self.resolve(part, d["Target"])
            out.append(d)
        return out

    def rel_target(self, part, rid):
        for r in self.rels(part):
            if r["Id"] == rid:
                return r.get("part")
        return None

    def add_rel(self, part, rel_type, target):
        name = self.rels_name(part)
        if not self.has(name):
            self.write(name, etree.Element(f"{{{PR_NS}}}Relationships", nsmap={None: PR_NS}))
        root = self.xml(name)
        used = {r.get("Id") for r in root}
        n = 1
        while f"rId{n}" in used:
            n += 1
        etree.SubElement(root, f"{{{PR_NS}}}Relationship", Id=f"rId{n}", Type=rel_type, Target=target)
        self.touch(name)
        return f"rId{n}"

    def remove_rel(self, part, rid):
        name = self.rels_name(part)
        root = self.xml(name)
        for r in list(root):
            if r.get("Id") == rid:
                root.remove(r)
        self.touch(name)

    # ---- content types ------------------------------------------------
    def set_override(self, part, content_type):
        root = self.xml("[Content_Types].xml")
        pn = "/" + part
        for o in root.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == pn:
                o.set("ContentType", content_type)
                break
        else:
            etree.SubElement(root, f"{{{CT_NS}}}Override", PartName=pn, ContentType=content_type)
        self.touch("[Content_Types].xml")

    def remove_override(self, part):
        root = self.xml("[Content_Types].xml")
        for o in root.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == "/" + part:
                root.remove(o)
        self.touch("[Content_Types].xml")

    def content_type(self, part):
        root = self.xml("[Content_Types].xml")
        for o in root.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == "/" + part:
                return o.get("ContentType")
        ext = part.rsplit(".", 1)[-1].lower()
        for d in root.findall(f"{{{CT_NS}}}Default"):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType")
        return None

    def is_template(self):
        ct = self.read("[Content_Types].xml")
        return any(k in ct for k in _MAIN_SWITCH)

    def as_document(self):
        """Turn a template (.dotx/.potx/.xltx) into a normal document in memory."""
        ct = self.read("[Content_Types].xml")
        for old, new in _MAIN_SWITCH.items():
            ct = ct.replace(old, new)
        self.write("[Content_Types].xml", ct)
        return self

    def next_part_name(self, pattern):
        """next_part_name('ppt/slides/slide{}.xml') -> first unused name."""
        n = 1
        while self.has(pattern.format(n)):
            n += 1
        return pattern.format(n)


# ---- canonical comparison / restore --------------------------------------
_UNORDERED = ("[Content_Types].xml",)


def canonical(name, data):
    """Byte form that ignores serialization noise (quotes, whitespace between
    tags, attribute order; child order for .rels / [Content_Types].xml)."""
    if not (name.endswith(".xml") or name.endswith(".rels") or name.endswith(".vml")):
        return data
    try:
        root = etree.fromstring(data, _BLANK_PARSER)
    except etree.XMLSyntaxError:
        return data
    if name.endswith(".rels") or name in _UNORDERED:
        return b"".join(sorted(etree.tostring(c, method="c14n") for c in root))
    return etree.tostring(root, method="c14n")


def restore_unchanged(original, result, dst=None):
    """After an engine rewrote the whole package (office_oxide, openpyxl, ...),
    copy back the original bytes of every part whose content is semantically
    unchanged and drop ZIP directory entries. Returns the list of parts that
    really changed."""
    with zipfile.ZipFile(original) as zo:
        orig = {i.filename: zo.read(i.filename) for i in zo.infolist() if not i.is_dir()}
    res = Package(result)
    changed = []
    for name in list(res._order):
        if name in orig and res._data[name] != orig[name]:
            if canonical(name, res._data[name]) == canonical(name, orig[name]):
                res.write(name, orig[name])
            else:
                changed.append(name)
    res.save(dst or result)
    return changed


def deepcopy(el):
    return copy.deepcopy(el)
