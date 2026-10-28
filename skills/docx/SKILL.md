---
name: docx
description: Create, edit and validate Word (.docx) documents. Use when the user asks for a Word document, report, letter, contract or any .docx output, to modify an existing .docx file, or to build a document on a user's Word template (.dotx/.docx).
---

# DOCX skill

## Tooling

| Need | Tool |
|---|---|
| Read any docx (incl. .dotx) | `python /opt/docgen/scripts/readback.py <file> [--format]` |
| Where does a text occur? | `python /opt/docgen/scripts/find_target.py <file> "<text>" [--section S] [--headings]` |
| **Edit an uploaded docx** (text, formatting, paragraphs, sections, tables, styles) | Python `docgen.docx` + `docgen.package.Package` (REFERENCE.md §docgen.docx) |
| Theme colors / fonts of a document | `docgen.theme.set_color` / `set_fonts` |
| New doc from markdown (fast path) | Python `office_oxide.create_from_markdown(md, "docx", out)` |
| New doc, full control (styles, sections, tables, headers, page numbers, per-run styling) | Node `docx` (docx-js) script in `/workspace/scratch/` |
| New doc on a **user template** | `docgen.docx.start_from_template` + python-docx |
| .doc / .rtf / .odt input | `convert_legacy.py` first, then edit the .docx |
| Render check (LibreOffice lays it out into pages) | `python /opt/docgen/scripts/render_check.py <file>` |
| Deliverable PDF of a docx | Gotenberg `/forms/libreoffice/convert` (`curl --fail-with-body`) |
| Validation | `validate.py <file> [--baseline <original>]`, `diff_check.py <original> <result> ...` |

`office_oxide.EditableDocument.replace_text` is NOT for targeted edits: it
replaces every occurrence, also inside longer words, and spreads formatting
across run boundaries.

## Creation workflow (new document)

1. Read the request. Decide: semantic structure (title, headings, sections,
   tables) and whether a theme or a user template applies.
2. **Fast path:** markdown in `/workspace/scratch/doc.md` (real heading
   levels, pipe tables), then
   `create_from_markdown(open(md).read(), "docx", "/workspace/outputs/<ad>.docx")`
   (it takes the markdown STRING; no per-element styling).
3. **Full-control path:** specific styling (custom styles, page numbers,
   headers/footers, complex tables, colors/fonts/sizes on particular text)
   → Node script with the `docx` package (REFERENCE.md §docx-js).
4. Verify: `validate.py` → `readback.py` → `render_check.py`.

## Editing workflow (uploaded document)

Never rebuild an uploaded document; change only what was asked.

1. `readback.py <file>` to see the structure; `find_target.py` to locate the
   target. Several matches → use the user's context (`section=`, table row,
   heading path). Still several → report the candidates, write no file.
2. Write `/workspace/scratch/edit.py` with the helpers:
   ```python
   from docgen.package import Package
   from docgen import docx as D, one
   src = "/workspace/uploads/rapor.docx"
   pkg = Package(src)

   # text: exactly one occurrence, formatting kept
   D.replace(one(D.find(pkg, "önerilir", section="Sonuç", whole_word=True), "önerilir"), "zorunludur")

   # character formatting of an expression (runs are split at the boundaries)
   m = one(D.find(pkg, "kritik sistemler"), "kritik sistemler")
   D.set_run_props(D.split_runs(m), pkg=pkg, color="006400")      # also bold/italic/underline/
                                                                  # strike/size/font/highlight; False removes
   # paragraph formatting of one paragraph
   D.set_paragraph_props([one(D.find_paragraphs(pkg, "Giriş", headings_only=True))], align="center")

   pkg.save("/workspace/outputs/rapor-duzenlenmis.docx")
   ```
   Scope rules:
   - one expression → `split_runs` + `set_run_props`; one paragraph → its runs /
     `set_paragraph_props`; **all headings / all body text** →
     `set_style_props(pkg, style_id, ...)` (`style_ids_for_level(pkg, 2)` gives
     the level-2 heading styles, incl. derived ones).
   - replace every occurrence only when asked: `D.replace_all(pkg, "risk", "tehdit",
     whole_word=True)`; inflected forms ("riski") stay — mention them.
   - sections: `D.section_blocks`, `D.append_to_section(pkg, heading, D.new_paragraph(text, like=para))`,
     `D.delete_section(pkg, heading)`; paragraphs: `D.set_paragraph_text`,
     `D.insert_after`, `D.delete`.
   - tables: `D.find_table`, `D.table_rows`, `D.set_cell_text`, `D.add_row`,
     `D.add_column`, `D.delete_column`, `D.set_cell_shading`.
   - theme recolor: `docgen.theme.set_color(pkg, "accent1", "2E7D32")`
     (author-set colors stay as they are). It also refreshes the hex values
     Word caches next to theme colors (incl. tints/shades), so render_check
     and PDF show the new colors; do not recompute them yourself. Verify with
     `diff_check.py SRC OUT --allow-theme --format-only`.
3. Verify:
   ```bash
   python /opt/docgen/scripts/validate.py OUT --baseline SRC
   python /opt/docgen/scripts/diff_check.py SRC OUT --allow-section "Sonuç" --text-only \
          --expect-text "zorunludur" --record /workspace/outputs/manifest.json
   python /opt/docgen/scripts/readback.py OUT --format     # style edits
   python /opt/docgen/scripts/render_check.py OUT
   ```
   Scope flags: `--allow-text`, `--allow-section`, `--allow-style ID`,
   `--allow-theme`, `--allow-part 'word/numbering.xml'`; `--text-only` for text
   edits, `--format-only` for formatting edits.

## User template (.dotx / .docx given as template)

The template's styles, theme, numbering, headers/footers (logo) and page
setup define the look; docgen themes do NOT apply.

```python
from docgen.docx import start_from_template, style_names
from docgen.package import Package
import docx as pydocx                      # python-docx

out = "/workspace/outputs/politika.docx"
start_from_template("/workspace/uploads/kurumsal.dotx", out)   # .dotx -> .docx, body emptied
print(style_names(Package(out)))           # use THESE style names
d = pydocx.Document(out)
d.add_paragraph("Uzaktan Çalışma Politikası", style="Heading 1")
d.add_paragraph("…", style="Normal")
d.save(out)
```
Verify with `validate.py out --baseline <template>`, readback (headers,
footers and logo must still be there) and `render_check.py out`. If the template has `{{PLACEHOLDER}}`
fields, fill them in place with `D.replace(one(D.find(pkg, "{{AD}}")), "…")`.

## Quality rules

- One `H1` (document title), `H2` for sections — never skip levels.
- Tables: caption above, sensible column widths, header row styled.
- Keep font consistency: default template typography unless a theme, a
  user template or the user specifies otherwise.
- Never deliver a file that fails `validate.py` (with `--baseline` for edits)
  or `diff_check.py`.

## Failure modes & fixes

| Symptom | Cause | Fix |
|---|---|---|
| `Ambiguous` from `one()` | text occurs several times | add `section=`, `whole_word=True`, table/row context; else report candidates |
| `diff_check` FAIL "character format (request was text-only)" | replacement spread formatting | replace only the exact expression (`find` + `replace`) |
| validate: error that also exists in the original | pre-existing Word quirk | always pass `--baseline <original>` for edits |
| style change not visible | run has direct formatting overriding the style, or theme font/color | change the runs, or remove the direct formatting (`False`) |
| theme recolor visible in Word but not in render_check / PDF | cached hex values of theme colors are stale (theme edited without `set_color`) | `docgen.theme.refresh_word_caches(pkg)` |
| `render_check` FAIL (Gotenberg HTTP error / no pages) | broken zip or malformed part | rebuild with `Package.save`, never hand-zip |
