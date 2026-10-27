# DOCX — technical reference

## office_oxide Python API (verified signatures)

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
- Text may be split across multiple `w:r`/`w:t` — `replace_text` (oxide)
  handles that; hand-editing does not (merge runs before replacing).

## Rezip rules (after unzipping + editing)

```bash
cd /workspace/scratch/doc_x
find . -name "*.DS_Store" -delete
zip -r -X ../new.docx . -x ".*"
```
- Zip from INSIDE the extracted dir so paths are relative (no top folder).
- No compression-level requirement for OOXML (unlike ODF `mimetype`).

## XSD validation (scripts/validate.py)

- `word/document.xml` → `schemas/wml.xsd`
- `word/styles.xml`, `word/numbering.xml` → `schemas/wml.xsd`
- Structural checks always run: zip integrity, `[Content_Types].xml`
  parses, all `.rels` parse, all XML parts well-formed.
