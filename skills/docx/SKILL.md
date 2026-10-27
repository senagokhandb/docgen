---
name: docx
description: Create, edit and validate Word (.docx) documents. Use when the user asks for a Word document, report, letter, contract or any .docx output, or to modify an existing .docx file.
---

# DOCX skill

## Tooling

| Need | Tool |
|---|---|
| New doc from markdown (fast path) | Python `office_oxide.create_from_markdown(md, "docx", out)` |
| Read any docx | Python `office_oxide.Document.open(p).to_markdown()` |
| Safe in-place text edit | Python `office_oxide.EditableDocument` (`replace_text`, `save`) |
| Full-control creation (styles, sections, tables, headers, page numbers) | Node `docx` (docx-js) script in `/workspace/scratch/` |
| Structural edit (styles.xml, headers, numbering) | unzip → edit XML with `lxml` → rezip |
| Style edit on existing doc (color, font, size, alignment) | unzip → edit `w:rPr` / `w:pPr` in `word/document.xml` with `lxml` → rezip (`EditableDocument` cannot style) |
| Render check / deliverable PDF | Gotenberg `/forms/libreoffice/convert` |
| Validation | `python /opt/docgen/scripts/validate.py <file>` |

## Creation workflow

1. Read this skill (done) and the user's request. Decide: semantic structure
   (title, headings, sections, tables) and whether a theme applies.
2. **Default path (fast):** write the content as markdown in
   `/workspace/scratch/doc.md` using real heading levels (`#` title, `##`
   sections) and pipe tables, then:
   ```python
   from office_oxide import create_from_markdown
   md = open("/workspace/scratch/doc.md").read()   # takes the STRING
   create_from_markdown(md, "docx", "/workspace/outputs/<ad>.docx")
   ```
   (`create_from_markdown` accepts the markdown text, not a file path.)
3. **Full-control path:** when the user asks for specific styling (custom
   styles, page numbers, headers/footers, complex tables, or a color /
   font / size / alignment on particular text), write a Node script using
   the `docx` package (see REFERENCE.md §docx-js) and run it with `node`.
   `create_from_markdown` cannot apply per-element styling.
4. Verify: `validate.py` → `readback.py` → Gotenberg office→PDF render.
   All three must pass.

## Editing workflow (existing file)

1. **Understand first:** `Document.open(path).to_markdown()` — read the
   current structure before touching anything.
2. **Simple text changes:**
   ```python
   from office_oxide import EditableDocument
   d = EditableDocument.open(path)
   n = d.replace_text("eski metin", "yeni metin")  # n = replacement count
   d.save(path)
   ```
   Unmodified OPC parts (styles, numbering, images) are preserved.
3. **Targeted style changes** ("bu başlığı kırmızı yap", "tablo başlıkları
   14 punto olsun"): unzip/rezip as in step 4, but patch only the matching
   runs/paragraphs in `word/document.xml` — do NOT rebuild the document.
   - Find the `w:r` runs whose `w:t` text matches; in each run's `w:rPr`
     (create it as the run's first child if missing) set:
     - font: `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>`
     - color: `<w:color w:val="FF0000"/>` (hex, no `#`)
     - size: `<w:sz w:val="28"/><w:szCs w:val="28"/>` (half-points: 28 = 14 pt)
     - child order is fixed by the schema: `w:rStyle`, `w:rFonts`, `w:b`,
       `w:i`, …, `w:color`, …, `w:sz`, `w:szCs`, `w:highlight`, `w:u` —
       insert in that order, replace existing ones instead of duplicating.
   - Alignment: `<w:jc w:val="center"/>` (`left`/`right`/`both`) in the
     paragraph's `w:pPr`.
   - Table cell background: `<w:shd w:val="clear" w:color="auto"
     w:fill="006400"/>` in that cell's `w:tcPr`.
   - If the target text is only part of a run, split the run (copy its
     `w:rPr` to the new runs) and style only the target part. Text may also
     be split across several runs — match on the paragraph's joined text.
   - "All headings" scope → change the style in `word/styles.xml`
     (e.g. `Heading1`) instead of every run.
   - An order mistake shows up in `validate.py` as "This element is not
     expected" — fix the order and re-run.
4. **Structural changes** (new styles, headers, numbering, section breaks):
   - unzip to `/workspace/scratch/` (`unzip -o file.docx -d scratch_docx`)
   - edit `word/document.xml` / `word/styles.xml` with `lxml`
     (mind the OOXML namespaces — see REFERENCE.md)
   - rezip from inside the folder: `cd scratch_docx && zip -r ../new.docx . -x ".*"`
5. Verify: validate → readback (confirm the exact change is present) →
   Gotenberg render.

## Quality rules

- One `H1` (document title), `H2` for sections — never skip levels.
- Tables: caption above, sensible column widths, header row styled.
- Keep font consistency: use the default template typography unless a theme
  or the user specifies otherwise.
- Never deliver a file that fails `validate.py`.
- For long documents: build section by section, read-back after each
  major section, not just at the end.

## Failure modes & fixes

| Symptom | Cause | Fix |
|---|---|---|
| validate.py: schema error on `w:` element | wrong namespace or missing `mc:Ignorable` | use canonical prefixes (see REFERENCE.md) |
| readback shows doubled/missing text | `replace_text` matched inside a split run | unzip and fix `w:t` runs directly |
| Gotenberg render fails | broken zip or malformed part | `zip -T` check, re-zip from clean extraction |
| styles not applied after XML edit | edited the wrong part (styles live in `word/styles.xml`, not `document.xml`) | re-locate the part |
