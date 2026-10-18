# xlsx-diff

A Python 3 command-line tool that compares two Excel `.xlsx` workbooks and writes a self-contained HTML report plus a CSV containing every reported change. It uses `openpyxl` and the Python standard library.

## Install on Linux

Requires Python 3.9 or newer with `venv` and `pip` available. From this project directory:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m pip install -e .
xlsx-diff --help
```

The editable install registers the `xlsx-diff` command. You can also run `python -m xlsx_diff` after installation.

## Usage

```bash
# Writes report.html and report.csv in the current directory.
xlsx-diff old.xlsx new.xlsx

# Choose the HTML filename; the CSV is results/comparison.csv.
mkdir -p results
xlsx-diff old.xlsx new.xlsx --output results/comparison.html

# Skip cell formatting differences.
xlsx-diff old.xlsx new.xlsx --ignore-formatting

# Skip cell comparisons where both sides are blank.
xlsx-diff old.xlsx new.xlsx --ignore-blank

# Compare cell values and formula expressions, plus sheet presence/order.
xlsx-diff old.xlsx new.xlsx --values-only --output values.html

# Options can be combined.
xlsx-diff old.xlsx new.xlsx --ignore-formatting --ignore-blank --output data.html

# Open the self-contained HTML report on a Linux desktop.
xdg-open data.html
```

The terminal prints added, removed and modified change counts, the number of affected worksheets, and both output paths. Exit status is `0` for a successful comparison (including when differences exist), `2` for invalid arguments, unreadable workbooks or output errors. Output directories must already exist. Existing report files are overwritten; destinations resolving to either input file, including symlinks and hard links, are rejected. Inputs are loaded and never saved.

## Comparison rules

- **Addresses:** `A1` is compared with `A1` on the same named sheet. Rows are never matched by their contents. **Inserting rows or columns can result in many reported cell differences.** Renamed sheets appear as a removed sheet and an added sheet.
- **Worksheets:** reports additions, removals and changes in the relative order of common worksheets. Adding or deleting a sheet alone does not count as reordering the remaining sheets. Order entries show one-based workbook positions. Added and removed sheets also have individual cell and layout changes reported, compared with an empty sheet.
- **Values:** reports additions, removals and modifications, retaining distinctions such as numbers versus text or booleans. Missing cells have a `None` value. Formula-to-literal changes are reported even when the displayed expression text is identical.
- **Formulas:** loads with `data_only=False` and compares the actual expressions. **Formulas are never evaluated**, cached calculation results are not compared, and changes in calculated results alone are not reported. Array/data-table formula objects are reported as their expression/attribute dictionaries.
- **Formatting:** compares fonts, fills, borders, alignment and number formats individually. Font/fill/border/alignment descriptions use openpyxl's XML representation, including colour identifiers and style attributes. Styled blank cells are included by default. Workbook theme definitions and rendered colour resolution are outside the comparison scope.
- **Layout:** compares merged ranges, hidden rows and columns, row heights, column widths, and worksheet default sizes/row visibility. Grouped column ranges are expanded into individual columns. Unspecified row heights inherit the sheet default. Unspecified column widths use the sheet's default width, or openpyxl's default of 13 when absent. Sizes are compared as stored, without rounding or attempting to render Excel's layout.
- **Sparse worksheets:** compares stored cells, including styled blank cells, without iterating every address in a bounding rectangle. Empty sheets and different dimensions are supported.

`--ignore-formatting` skips the five cell formatting categories; merged ranges, hidden rows/columns and sizes are still compared. `--ignore-blank` skips a cell only when **both** values are `None` or an empty string, including that cell's formatting. It does not suppress a nonempty value added to or deleted from a blank cell, or worksheet layout changes. `--values-only` skips all formatting and layout comparisons, retaining cell values/formulas and sheet presence/order.

Only the listed properties are compared. Comments, hyperlinks, charts, images, validations, conditional formatting, protection, print settings and other workbook features are outside scope. Workbook formula dependencies and references are not interpreted.

## Reports

The HTML report includes totals, a worksheet navigation list, and tables grouped by worksheet. Added entries are green, removed entries red, and modified entries yellow; each row also has an explicit change label. All workbook content is HTML-escaped. No scripts, stylesheets, fonts or other external resources are required.

The UTF-8 CSV has these columns:

| Column | Meaning |
| --- | --- |
| `sheet` | Worksheet name |
| `reference` | Cell, merged range, row (`5:5`), column (`E:E`), or blank for worksheet properties |
| `change_type` | `added`, `removed`, or `modified` |
| `property` | `value`, `formula`, formatting component, layout property, `worksheet`, or `worksheet_order` |
| `old_value` | Previous value; empty for absent values |
| `new_value` | Updated value; empty for absent values |

Counts are **individual property changes**, not unique cells: a value and a font change on one cell count as two changes. Layout and formatting changes use `modified`; additions/removals apply to cell values, worksheet presence and merged ranges. CSV values preserve the original text, including formula expressions; use a CSV parser when consuming them programmatically.

## Run tests

```bash
source .venv/bin/activate
python -m unittest discover -s tests -v
```

Tests create temporary workbooks and cover cell types, formula expressions, all requested formatting categories, layout, missing/reordered/empty sheets, sparse dimensions, HTML escaping, CSV quoting, CLI errors, output aliases and unchanged input hashes.
