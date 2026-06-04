# org-to-xlsx

Extract every table from an Org-mode file into a multi-sheet Excel workbook —
one sheet per table, with the surrounding paragraph carried over as a note
block above the data.

Built for sharing Org-mode estimates and reports with people who live in Excel
but don't read `.org`.

## Why

Org-mode tables are great for the author (formulas, alignment, version
control), but stakeholders usually want `.xlsx`. The standard
`org-table-export` produces a single CSV per table; this tool produces a
single workbook for the whole file, preserves heading structure, and keeps the
explanatory text next to the data it describes.

## Features

- **One sheet per table.** Sheet name is derived from the nearest enclosing
  Org heading, sanitized for Excel (no `[]:*?/\`, max 31 chars), with
  automatic `(2)`, `(3)` suffixes for collisions.
- **Heading breadcrumbs.** Each sheet shows the full heading path (e.g.
  `Module Breakdown / CMS Knowledge Core`) under the title so context isn't
  lost.
- **Description note block.** Any paragraph(s) that appear between the heading
  and the table become a merged italic note above the table — the same
  context the reader would see in the Org file.
- **Org markup cleanup.** Strips `*bold*`, `/italic/`, `=verbatim=`, `~code~`,
  `_underline_` wrappers when they enclose a whole cell, so values land in
  Excel as plain text.
- **Handles `#+TBLFM:` footers** and separator rows (`|---|`) correctly.
- **Sensible formatting.** Bold header row with light-blue fill, borders,
  text wrap, frozen header pane, auto-sized columns capped at 60 chars.
- **Zero configuration.** Point it at an `.org` file, get an `.xlsx` back.

## Requirements

- Python 3.11+
- [uv](https://github.com/astral-sh/uv) (recommended) or any PEP 517 installer

Runtime dependencies (installed automatically):

- [`orgparse`](https://github.com/karlicoss/orgparse) — Org-mode parser
- [`xlsxwriter`](https://github.com/jmcnamara/XlsxWriter) — XLSX generator

## Installation

### As a CLI tool (recommended)

```bash
uv tool install git+https://github.com/<you>/org-to-xlsx
org-to-xlsx path/to/file.org
```

After this, `org-to-xlsx` is available globally on your `PATH`.

### From a clone

```bash
git clone https://github.com/<you>/org-to-xlsx
cd org-to-xlsx
uv sync
uv run org-to-xlsx path/to/file.org
```

### With `pip`

```bash
pip install git+https://github.com/<you>/org-to-xlsx
```

## Usage

```bash
org-to-xlsx INPUT.org [-o OUTPUT.xlsx]
```

Arguments:

| Arg | Description |
|---|---|
| `INPUT` | Source `.org` file (required) |
| `-o`, `--output` | Output `.xlsx` path. Defaults to `INPUT.xlsx` next to the source. |

Examples:

```bash
# Default: writes estimate.xlsx next to estimate.org
org-to-xlsx estimate.org

# Explicit output path
org-to-xlsx estimate.org -o ~/Desktop/Estimate.xlsx

# Run without installing (uv handles deps from the cloned repo)
uv run --project /path/to/org-to-xlsx org-to-xlsx file.org
```

Sample output:

```
$ org-to-xlsx estimate.org
Wrote 15 table(s) to estimate.xlsx
```

## How it parses

For each top-level and nested heading in the Org file:

1. The body text is scanned line by line.
2. Consecutive lines starting with `|` form a table; preceding non-blank
   lines (since the previous table or the heading) are captured as the
   table's description.
3. Separator rows (`|---|`) are dropped. `#+TBLFM:` footers are skipped.
4. Each captured table becomes a worksheet:
   - Row 0: heading title (bold, 14pt)
   - Row 1: full heading breadcrumb (italic, gray)
   - Row 3: description block (merged italic across the table width), if any
   - Row 5+: header row + data rows

A heading with multiple tables produces multiple sheets, with `(2)`, `(3)`,
... suffixes appended to keep names unique.

## Limitations & non-goals

- **No formulas.** Org `#+TBLFM:` lines are ignored — the tool exports
  computed values as they appear in the source. Re-evaluate formulas in Org
  (`C-c C-c` on the table) before exporting if you need fresh totals.
- **No styles inside cells.** Bold/italic markers wrapping a whole cell are
  stripped, but inline emphasis inside cells (`some *bold* word`) is left
  as-is. This is intentional — full Org markup parsing inside table cells
  would balloon the scope.
- **No images, no source blocks, no drawers.** Only tables are extracted.
- **Sheet names are truncated to 31 chars** (Excel limit). Use distinct
  heading prefixes if collisions matter.
- **No CSV / ODS export.** XLSX only, by design.

## Development

```bash
git clone https://github.com/<you>/org-to-xlsx
cd org-to-xlsx
uv sync
uv run org-to-xlsx --help
```

The whole tool is one file: `src/org_to_xlsx/__init__.py`. Patches welcome.

## License

MIT
