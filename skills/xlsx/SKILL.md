---
name: xlsx
description: Create and edit Excel (.xlsx) workbooks with data, formulas and formatting. Use for budgets, financial models, data tables, calculations, or any .xlsx task.
---

# XLSX skill

## Tooling

| Need | Tool |
|---|---|
| New workbook with formulas + basic styles (bold + background only) | Python `office_oxide.XlsxWriter` |
| Precise cells, styles (font color/name/size, fill, alignment), charts, conditional formats | Python `openpyxl` — `Font(name=, size=, color=)`, `PatternFill`, `Alignment(horizontal=)` per cell or range |
| Data wrangling / large imports | Python `pandas` (optional) |
| Read any xlsx (fast overview) | Python `office_oxide.Document.open(p).to_markdown()` |
| Recalculation QA (formulas!) | `python /opt/docgen/scripts/recalc_check.py <file>` |
| Render check | Gotenberg `/forms/libreoffice/convert` (xlsx→PDF) |
| Validation | `python /opt/docgen/scripts/validate.py <file>` |

> **Important:** written formula cells carry **no cached value**. Excel and
> LibreOffice compute them on open — which is why formula QA is a dedicated
> step (`recalc_check.py`), not a visual glance.

## Creation workflow

1. Plan the layout before coding: sheet list, column headers, units, and
   which cells are **inputs** vs **formulas** vs **cross-sheet links**.
2. Write a Python script in `/workspace/scratch/build_xlsx.py`.

   **XlsxWriter path (default):**
   ```python
   from office_oxide import XlsxWriter
   wb = XlsxWriter()
   s = wb.add_sheet("Bütçe")                    # -> sheet index 0
   wb.set_column_width(s, 0, 32)                # col A (0-based)
   wb.set_cell_styled(s, 0, 0, "Kalem", bold=True)   # A1  (row/col 0-based)
   wb.set_cell_styled(s, 0, 1, "Tutar", bold=True)   # B1
   wb.set_cell(s, 1, 0, "Malzeme")
   wb.set_cell(s, 1, 1, 12000)                  # input  (B2)
   wb.set_cell(s, 2, 0, "İşçilik")
   wb.set_cell(s, 2, 1, 8000)                   # input  (B3)
   wb.set_formula(s, 3, 1, "=SUM(B2:B3)")       # B4 — formula WITH "="
   wb.merge_cells(s, 0, 0, 1, 2)                # merge A1:B1
   wb.save("/workspace/outputs/butce.xlsx")
   ```
   **API gotchas:**
   - `row`/`col` are **0-based integers** (A1 = row 0, col 0).
   - Formulas in `set_formula` **include the leading `=`**; references
     inside the formula use normal A1 notation.
   - `EditableDocument.set_cell(sheet_index, cell_ref, value)` is different:
     it takes an **A1 string** (`"B4"`) — use it for edits on existing files.

   **openpyxl path:** use it when you need charts, data validation,
   conditional formatting, or to load+modify an existing workbook.

3. Save the final file to `/workspace/outputs/`.
4. **QA — all three, in order:**
   1. `python /opt/docgen/scripts/validate.py <file.xlsx>`
   2. `python /opt/docgen/scripts/readback.py <file.xlsx>`
   3. `python /opt/docgen/scripts/recalc_check.py <file.xlsx>`
      → converts via LibreOffice (Gotenberg), reads the rendered PDF and
      scans for error tokens (`#REF!`, `#NAME?`, `#DIV/0!`, `#VALUE!`,
      `#N/A`, `#NULL!`, `#NUM!`). Must report 0 errors.

## Editing workflow

1. Read first: `Document.open().to_markdown()` for a fast structural
   overview; `openpyxl` (`data_only=False`) when you need the actual
   formulas.
2. Modify with `openpyxl` (`load_workbook` → edit → `save`) — it preserves
   most workbook parts.
3. QA: the same three steps. If the file has formulas, `recalc_check.py`
   is mandatory, not optional.

## Formula rules (critical)

**Supported set:** standard Excel 2007+ core functions.

**BANNED** (not supported by all engines — use alternatives):
| Banned | Use instead |
|---|---|
| `XLOOKUP` | `INDEX` + `MATCH` |
| `XMATCH` | `MATCH` |
| `SORT` | helper column + `RANK` / manual order |
| `FILTER` | `SUMIFS` / `COUNTIFS` / helper columns |
| `UNIQUE` | helper column + `COUNTIF` |
| `SEQUENCE` | literal range or `ROW()` trick |

**`_xlfn.` prefix required** for these newer functions:
`_xlfn.TEXTJOIN`, `_xlfn.CONCAT`, `_xlfn.IFS`, `_xlfn.SWITCH`,
`_xlfn.MAXIFS`, `_xlfn.MINIFS`.

**Never** hardcode a number the user expects to be a live calculation.

## Financial color convention (apply to models/budgets)

| Meaning | Style |
|---|---|
| Manual input | blue font (`#0000FF`) |
| Formula | black font |
| Link to another sheet (internal) | green font (`#008000`) |
| External reference | red font (`#FF0000`) |
| Key assumption | yellow fill (`#FFFF00`) |

This is the default convention. If the user explicitly asks for a color,
font or size on specific cells, apply exactly that to those cells
(QWEN.md › Themes); keep the convention everywhere else.

## Verification

- `recalc_check.py` MUST pass for any workbook containing formulas.
- For critical totals, double-check by computing the expected value in
  Python and comparing against the read-back.
- Max 3 fix iterations on a failing QA step, then report honestly.
