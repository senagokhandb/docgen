# DOCX — technical reference

## docgen.docx — in-place editing API (tested on Word-made files)

```python
from docgen.package import Package          # open / save keeping untouched parts byte-identical
from docgen import docx as D, one           # one(list, "what") -> the single item, else Ambiguous/NotFound
pkg = Package(path)                         # .docx or .dotx
pkg.save(dst)                               # only modified parts are re-serialized

# --- locate -----------------------------------------------------------
D.paragraphs(pkg)            # all paragraphs: .text .style .level .path (heading chain) .table (t,r,c) .where()
D.headings(pkg)              # heading paragraphs (Title=0, Heading N=N, derived styles included)
D.find(pkg, text, section="Bulgular > Sonuç", whole_word=True, ignore_case=False,
       part="document|header|footer", in_table=True|False, headings_only=False)   # -> [Match]
D.find_paragraphs(pkg, contains, section=None, style=None, headings_only=False)    # -> [Para]

# --- text --------------------------------------------------------------
D.replace(match, "yeni")                    # one occurrence; returns a warning if it spanned mixed formatting
D.replace_all(pkg, old, new, whole_word=True, **find_filters)   # -> (count, warnings)
D.set_paragraph_text(para, "tüm paragraf metni")                 # keeps pPr + first run format
D.new_paragraph(text, like=para, style=None)                     # element matching an existing paragraph
D.insert_after(para_or_element, *elements)
D.section_blocks(pkg, heading)  /  D.append_to_section(pkg, heading, *elements)
D.delete([para_or_element, ...])  /  D.delete_section(pkg, heading)

# --- formatting --------------------------------------------------------
runs = D.split_runs(match)                  # exact runs for the match (also split_runs(para, start, end))
D.set_run_props(runs, pkg=pkg, bold=True, italic=None, underline=True|"double"|False,
                strike=False, color="C00000", size=11, font="Times New Roman", highlight="yellow")
#   None = keep, False = remove (switched off explicitly when a style turns it on)
D.set_paragraph_props(paras, align="left|center|right|justify", space_before=6, space_after=12,
                      line_spacing=1.5, indent_left_cm=1, indent_first_cm=0.5, style="Heading2")
D.set_style_props(pkg, "Heading2", color="1F3864", size=14, bold=True, align="center")  # every paragraph of that style
D.style_ids_for_level(pkg, 2)               # ['Heading2', ...derived]
D.style_names(pkg)                          # {styleId: name}

# --- tables ------------------------------------------------------------
t = one(D.find_table(pkg, "Bölge"))         # tables whose text contains ...
D.table_rows(t)                             # [[cell text]]
D.set_cell_text(t, row, col, "3.250.000")   # keeps the cell's formatting
D.add_row(t, after=2, values=[...]);  D.add_column(t, after=1, values=[...]);  D.delete_column(t, 2)
D.set_cell_shading(t, 0, 1, "1F3864")

# --- templates ---------------------------------------------------------
D.start_from_template(template, dst, keep_content=False)   # .dotx/.docx -> .docx, body emptied
```
Errors: `docgen.NotFound`, `docgen.Ambiguous` (`.candidates`), `docgen.NotSupported`
(e.g. merged cells for add_column). Theme: `docgen.theme.set_color(pkg, "accent1", "2E7D32")`
(also refreshes the cached hex values of theme-bound colors; `refresh_word_caches(pkg)` does
only that), `set_fonts(pkg, major="Georgia", minor="Calibri")`.

## office_oxide Python API (verified signatures)

`EditableDocument.replace_text` replaces EVERY occurrence, also inside longer
words, and can spread formatting across runs — use `docgen.docx` for edits.

```python
from office_oxide import (Document, EditableDocument,
                          create_from_markdown, OfficeOxideError)

# read
doc = Document.open("rapor.docx")          # or Document.from_bytes(data, "docx")
doc.to_markdown()                          # -> str
doc.plain_text()
doc.to_markdown_with_images()
doc.to_html()
doc.to_ir() / doc.to_ir_json()             # structured intermediate repr
doc.save_as("other.docx")                  # also converts between 6 formats:
                                           # docx/xlsx/pptx + legacy doc/xls/ppt
doc.close()
# (Document is a context manager: `with Document.open(p) as d: ...`)

# safe in-place edit (preserves unmodified OPC parts)
d = EditableDocument.open("rapor.docx")
n = d.replace_text("eski", "yeni")         # n = replacement count
d.set_cell(0, "B4", 123.5)                 # xlsx only: (sheet_index, "A1"-ref, value)
d.save("rapor.docx")

# create from markdown (takes the STRING)
create_from_markdown(md_text, "docx", "out.docx")
```

## docx package (docx-js) — key API

```js
const { Document, Packer, Paragraph, TextRun, HeadingLevel,
        Table, TableRow, TableCell, WidthType, AlignmentType,
        Header, Footer, PageNumber, NumberFormat } = require("docx");

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph("Başlık")] }) },
    footers: { default: new Footer({ children: [
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }) },
    children: [
      new Paragraph({ text: "Başlık", heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ text: "Giriş", heading: HeadingLevel.HEADING_2 }),
      new Paragraph("Metin..."),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [ new TableRow({ children: [
          new TableCell({ children: [new Paragraph("H1")] }),
          new TableCell({ children: [new Paragraph("H2")] }),
        ]}) ],
      }),
    ],
  }],
});
Packer.toBuffer(doc).then(buf => require("fs").writeFileSync("out.docx", buf));
```

Notes:
- Sizes are half-points (22 = 11pt). Margins in twips (1440 = 1 inch).
- Per-element styling (user-requested color/font/size/alignment):
  ```js
  new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new TextRun({ text: "Normal " }),
    new TextRun({ text: "kırmızı", color: "FF0000", font: "Arial", size: 28, bold: true }),
  ] });
  new TableCell({ shading: { type: ShadingType.CLEAR, color: "auto", fill: "006400" },
    children: [new Paragraph({ children: [new TextRun({ text: "Başlık", color: "FFFFFF" })] })] });
  ```
  (`ShadingType` is also imported from `docx`; colors are hex without `#`.)
- `Packer.toBuffer` + `fs.write` keeps everything in the container.

## OOXML anatomy (docx = zip)

```
[Content_Types].xml        part → content-type map
_rels/.rels                package relationships
word/document.xml          main body (w: namespace)
word/styles.xml            named styles (Heading1, Title, ...)
word/numbering.xml         lists
word/headerN.xml / footerN.xml
word/media/                images
word/_rels/document.xml.rels
```

Namespaces (canonical prefixes — keep them):
- `w`  = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- `r`  = `.../officeDocument/2006/relationships`
- `mc` = `.../markup-compatibility/2006` (root carries `mc:Ignorable="w14 w15 ..."`;
  keep that attribute when editing)
- `wp`, `a`, `pic` for drawings/images

Run of text: `<w:p><w:r><w:t>metin</w:t></w:r></w:p>`
- Text may be split across multiple `w:r`/`w:t` — `docgen.docx.find` matches
  on the joined paragraph text and `split_runs` cuts runs at exact offsets.

## Saving edited packages

Use `docgen.package.Package(...).save(dst)` instead of unzip/zip: it keeps
every untouched part byte-identical, keeps the entry order and adds no ZIP
directory entries (`zip -r` adds them; if you ever must hand-zip, use
`zip -r -D -X ../new.docx .` from inside the extracted folder).
If another engine rewrote the whole package, run
`docgen.package.restore_unchanged(original, result)` to put back the original
bytes of semantically unchanged parts.

## XSD validation (scripts/validate.py)

- `word/document.xml` → `schemas/wml.xsd`
- `word/styles.xml`, `word/numbering.xml` → `schemas/wml.xsd`
- Structural checks always run: zip integrity, `[Content_Types].xml`
  parses, all `.rels` parse, all XML parts well-formed.
- Word itself writes a few things the 1st-edition schema rejects (e.g.
  `<w:zoom w:val="bestFit"/>` without `w:percent`). For edits and template
  builds pass `--baseline <original>`: findings already in the original are
  listed as pre-existing and do not fail the check.
