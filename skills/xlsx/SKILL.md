---
name: xlsx
description: Create and edit Excel (.xlsx) workbooks with data, formulas and formatting. Use for budgets, financial models, data tables, calculations, editing an uploaded .xlsx, or filling a user's Excel template (.xltx/.xlsx).
---

# XLSX skill

## Tooling

| Need | Tool |
|---|---|
| Read a workbook (values / formats) | `readback.py <file>` / `readback.py <file> --format` |
| Where is a value? | `find_target.py <file> "<text>" [--sheet S]` |
| **Edit an uploaded workbook**: values, formulas, cell formats, number formats, conditional formats, column widths | Python `docgen.xlsx` + `docgen.package.Package` (REFERENCE.md §docgen.xlsx) |
| Insert/delete rows or columns, sort a range (references shift everywhere) | `python /opt/docgen/scripts/xlsx_struct.py ...` |
| Correct cached formula results (after ANY value/formula change) | `python /opt/docgen/scripts/xlsx_recalc.py <file> [--expect "Sheet!B4=20000"]` |
| Theme colors / fonts | `docgen.theme.set_color` / `set_fonts` |
| New workbook with formulas + basic styles (bold + background only) | Python `office_oxide.XlsxWriter` |
| New workbook with full styling, charts, validation | Python `openpyxl` (new files only) |
| Fill a **user template** (.xltx/.xlsx) | `Package(template).as_document()` + `docgen.xlsx` |
| .xls / .ods input | `convert_legacy.py` first |
| Render check (LibreOffice lays it out into pages) | `python /opt/docgen/scripts/render_check.py <file>` |
| Validation | `validate.py <file> [--baseline <original>]`, `diff_check.py <original> <result> ...` |

> openpyxl `load_workbook` → `save` on an UPLOADED file drops every cached
> formula result and rewrites external links, drawings and styles. Use it
> only for new workbooks; edit uploads with `docgen.xlsx`.
> office_oxide `EditableDocument.set_cell` leaves dependent results stale.

## Creation workflow (new workbook)

1. Plan: sheet list, headers, units; which cells are **inputs** vs
   **formulas** vs **cross-sheet links**.
2. `/workspace/scratch/build_xlsx.py` with XlsxWriter (simple) or openpyxl
   (styles/charts/validation):
   ```python
   from office_oxide import XlsxWriter
   wb = XlsxWriter()
   s = wb.add_sheet("Bütçe")                    # -> sheet index 0
   wb.set_column_width(s, 0, 32)                # col A (0-based)
   wb.set_cell_styled(s, 0, 0, "Kalem", bold=True)   # A1  (row/col 0-based)
   wb.set_cell(s, 1, 0, "Malzeme"); wb.set_cell(s, 1, 1, 12000)
   wb.set_formula(s, 3, 1, "=SUM(B2:B3)")       # B4 — formula WITH "="
   wb.save("/workspace/outputs/butce.xlsx")
   ```
3. QA: `validate.py` → `xlsx_recalc.py <file>` (fills results, reports
   `#REF!`/`#DIV/0!`…; must PASS) → `readback.py` → `render_check.py`.

## Editing workflow (uploaded workbook)

Never rebuild an uploaded workbook. Tables, conditional formats,
validations, charts, images, names and external links must survive.

1. `readback.py` / `find_target.py` to locate. Rows are found by a key
   column: `X.rows_where(pkg, "Durum", "Servis", "Servis B")`. Several rows
   match → report them, write no file.
2. Edit:
   ```python
   from docgen.package import Package
   from docgen import xlsx as X, one
   pkg = Package("/workspace/uploads/satis.xlsx")
   X.set_value(pkg, "Satışlar!E2", 18, expect=16)        # expect = the value the user named
   X.set_style(pkg, "Durum!A1:C1", font={"color": "FFFFFF"}, fill="1F3864")
   X.set_style(pkg, "Satışlar!G2:G41", number_format="#,##0.00")
   X.add_conditional_format(pkg, "Satışlar!E2:E41", operator="greaterThan", formula="20", fill="C6EFCE")
   pkg.save("/workspace/outputs/satis-duzenlenmis.xlsx")
   ```
   - `set_style` changes only the named properties; `False` removes one
     (`font={"underline": False}`) — never reset the whole format.
   - Guards: overwriting a formula → `NeedsConfirmation` (needs `--onay`,
     then `allow_formula_overwrite=True`); `expect` mismatch → `ValueMismatch`
     (report the real value); values starting with `=` → use `set_formula`.
   - Rows/columns/sort: `xlsx_struct.py IN OUT insert-rows --sheet S --at 10`,
     `delete-rows` (refuses when formulas/names read the rows; `--onay` only
     if the user confirmed), `sort --sheet S --range A1:H41 --by G --desc`.
3. **Recalculate** (mandatory after values/formulas; `xlsx_struct.py` does it
   itself): `python /opt/docgen/scripts/xlsx_recalc.py OUT`.
4. Verify:
   ```bash
   python /opt/docgen/scripts/validate.py OUT --baseline SRC
   python /opt/docgen/scripts/diff_check.py SRC OUT --allow-cell "Satışlar!E2" --allow-recalc \
          --record /workspace/outputs/manifest.json
   python /opt/docgen/scripts/render_check.py OUT
   ```
   Scope flags: `--allow-cell 'Sheet!A1:C9'`, `--allow-sheet S`,
   `--allow-recalc` (results of dependent formulas), `--allow-structure`
   (row/column shifts), `--allow-feature cf:Sheet|dv|merged|cols|tables|names|charts`.
   Example: KDV rate change → `--allow-cell "Parametreler!B2" --allow-recalc`;
   a summary sheet that does not depend on it must NOT show changes.

## User template (.xltx / .xlsx given as template)

```python
from docgen.package import Package
from docgen import xlsx as X
pkg = Package("/workspace/uploads/teklif.xltx").as_document()   # template -> workbook
X.set_value(pkg, "Teklif!B3", "ACME Ltd")
pkg.save("/workspace/outputs/teklif-acme.xlsx")
```
Logo, merged cells, styles and formulas of the template stay as they are;
then `xlsx_recalc.py`, `validate.py --baseline <template>` and `render_check.py`.

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

## Financial color convention (apply to new models/budgets)

| Meaning | Style |
|---|---|
| Manual input | blue font (`#0000FF`) |
| Formula | black font |
| Link to another sheet (internal) | green font (`#008000`) |
| External reference | red font (`#FF0000`) |
| Key assumption | yellow fill (`#FFFF00`) |

This is the default convention for new workbooks. If the user explicitly
asks for a color, font or size on specific cells, apply exactly that to
those cells (QWEN.md › Themes); keep the convention everywhere else. Do not
impose it on uploaded workbooks.

## Verification

- `xlsx_recalc.py` MUST pass for any workbook containing formulas.
- Edits: `diff_check.py` must PASS with the declared scope.
- Max 3 fix iterations on a failing QA step, then report honestly.
