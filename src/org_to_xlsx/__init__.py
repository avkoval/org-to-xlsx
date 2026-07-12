"""Extract org-mode tables from an .org file into an .xlsx workbook.

Each table becomes its own sheet. The sheet name is derived from the
nearest enclosing heading. Any text that appears between the heading and
the table is written above the table as an italic note block, so the
context for each table travels with it into Excel.

Optional consolidation mode
---------------------------
If the source .org file contains a directive like

    #+CONSOLIDATE_TOTALS: Dev total

the tool additionally prepends a "Summary" sheet that lists, per chapter
(`** N. CODE — Name` heading), the numeric value from the row in the
chapter's task table whose label (first non-empty cell) contains the
directive value ("Dev total" above). The Summary sheet also contains an
editable MIN column that users can fill in with their own estimate; the
column is preserved verbatim when the file is regenerated.

Without the directive the tool behaves exactly as before.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

import orgparse
import xlsxwriter
from orgparse.node import OrgNode

__all__ = ["main"]

TABLE_LINE_RE = re.compile(r"^\s*\|")
SEPARATOR_LINE_RE = re.compile(r"^\s*\|[-+]+\|?\s*$")
TBLFM_LINE_RE = re.compile(r"^\s*#\+TBLFM:", re.IGNORECASE)
INVALID_SHEET_CHARS_RE = re.compile(r"[\[\]:*?/\\]")
WHITESPACE_RE = re.compile(r"\s+")
EMPHASIS_CHARS = {"*", "/", "_", "=", "~"}

# --- Consolidation mode helpers -------------------------------------------

CONSOLIDATE_DIRECTIVE_RE = re.compile(
    r"^\s*#\+CONSOLIDATE_TOTALS:\s*(.+?)\s*$", re.IGNORECASE
)
# Match `** N. Rest` — Rest may contain "CODE — Name" or just plain text.
CHAPTER_HEADING_RE = re.compile(r"^\*\*\s+(\d+)\.\s*(.+?)\s*$", re.UNICODE)
# Match `* Part X. Rest` at document level.
PART_HEADING_RE = re.compile(
    r"^\*\s+Part\s+([0-9IVX]+)\.\s*(.+?)\s*$", re.UNICODE
)
NUMERIC_RE = re.compile(r"(\d+)")

# Well-known multi-word chapter codes we don't want to split at the hyphen.
MULTIWORD_CHAPTER_CODES = (
    "CAB-CAND", "CAB-EXAM", "CAB-AI", "CAB-COM",
    "EXAM-RUN", "NFR-REL", "NFR-PERF", "NFR-A11Y", "NFR-COMP",
    "DEV-ENV",
)
CODE_SPLIT_RE = re.compile(
    r"^\s*([A-Za-zА-Яа-я0-9\-]+(?:\s+[A-Za-zА-Яа-я0-9\-]+)?)\s+[—–\-]\s+",
    re.UNICODE,
)


@dataclass
class ExtractedTable:
    heading: str
    full_path: str
    description: list[str]
    rows: list[list[str]]


@dataclass
class ChapterTotal:
    """One row of the consolidation summary sheet."""

    number: int
    part: str
    code: str
    name: str
    total: float | None
    total_raw: str = ""
    notes: list[str] = field(default_factory=list)


# --- Consolidation parsing ------------------------------------------------


def read_directive(org_path: Path) -> str | None:
    """Return the value of `#+CONSOLIDATE_TOTALS:` if the directive is present."""
    for line in org_path.read_text(encoding="utf-8").splitlines():
        m = CONSOLIDATE_DIRECTIVE_RE.match(line)
        if m:
            return m.group(1).strip()
    return None


def strip_org_wrappers(text: str) -> str:
    """Drop emphasis markers around a whole cell so 'Dev total' matches '*Dev total*'."""
    text = text.strip()
    while len(text) >= 2 and text[0] == text[-1] and text[0] in EMPHASIS_CHARS:
        text = text[1:-1].strip()
    return text


def extract_chapter_code(rest: str) -> tuple[str, str]:
    """Split '`CODE — Description`' remainder into (code, remaining name)."""
    text = rest.strip()
    for prefix in MULTIWORD_CHAPTER_CODES:
        if text.upper().startswith(prefix):
            leftover = text[len(prefix):]
            return prefix, leftover.lstrip(" —–-")
    m = CODE_SPLIT_RE.match(text)
    if m:
        return m.group(1).strip().upper(), text[m.end():].strip()
    # Fall back to the first whitespace-delimited token.
    parts = text.split(" ", 1)
    return parts[0].upper(), parts[1] if len(parts) > 1 else ""


def parse_totals(org_path: Path, marker: str) -> list[ChapterTotal]:
    """Walk the .org file and collect one `ChapterTotal` per `** N.` chapter.

    Within each chapter, find the *first* table row whose first non-empty cell
    (after stripping emphasis) contains ``marker`` (case-insensitive). Extract
    the first integer from the row's next cell as the chapter's total. Ranges
    like ``20-40`` return the low end; both raw and parsed values are kept.
    """
    lines = org_path.read_text(encoding="utf-8").splitlines()

    marker_lc = marker.lower()
    chapters: list[ChapterTotal] = []
    current_part = "?"
    current: ChapterTotal | None = None

    for line in lines:
        m = PART_HEADING_RE.match(line)
        if m:
            current_part = m.group(1).strip()
            continue

        m = CHAPTER_HEADING_RE.match(line)
        if m:
            number = int(m.group(1))
            code, remaining = extract_chapter_code(m.group(2))
            current = ChapterTotal(
                number=number,
                part=current_part,
                code=code,
                name=remaining,
                total=None,
            )
            chapters.append(current)
            continue

        if current is None or current.total is not None:
            continue
        if not TABLE_LINE_RE.match(line):
            continue
        if SEPARATOR_LINE_RE.match(line):
            continue

        row = line.strip()
        if row.startswith("|"):
            row = row[1:]
        if row.endswith("|"):
            row = row[:-1]
        cells = [strip_org_wrappers(c) for c in row.split("|")]

        # Look through cells for the marker; the numeric total is the *next*
        # non-empty cell after the marker cell (works for `| / | *X — Dev total* | 29 |`).
        for i, cell in enumerate(cells):
            if marker_lc in cell.lower():
                for follow in cells[i + 1:]:
                    if follow.strip():
                        num = NUMERIC_RE.search(follow)
                        if num:
                            current.total = float(num.group(1))
                            current.total_raw = follow.strip()
                        break
                break

    return chapters


def strip_emphasis(cell: str) -> str:
    """Drop org emphasis markers (*bold*, /italic/, =verbatim=) wrapping a cell."""
    cell = cell.strip()
    if len(cell) >= 2 and cell[0] == cell[-1] and cell[0] in EMPHASIS_CHARS:
        return cell[1:-1].strip()
    return cell


def parse_table(lines: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in lines:
        if SEPARATOR_LINE_RE.match(line):
            continue
        stripped = line.strip()
        if stripped.startswith("|"):
            stripped = stripped[1:]
        if stripped.endswith("|"):
            stripped = stripped[:-1]
        cells = [strip_emphasis(cell) for cell in stripped.split("|")]
        rows.append(cells)
    return rows


def extract_from_body(body: str) -> Iterator[tuple[list[str], list[list[str]]]]:
    """Yield (description_lines, table_rows) tuples found in a heading body."""
    lines = body.splitlines()
    pending_description: list[str] = []
    table_lines: list[str] = []
    description_for_current: list[str] = []
    in_table = False

    for line in lines:
        if TABLE_LINE_RE.match(line):
            if not in_table:
                in_table = True
                description_for_current = list(pending_description)
                pending_description = []
                table_lines = []
            table_lines.append(line)
            continue

        if in_table:
            # A non-pipe line ends the current table; ignore #+TBLFM: footers.
            if TBLFM_LINE_RE.match(line):
                continue
            yield description_for_current, parse_table(table_lines)
            in_table = False
            table_lines = []
            description_for_current = []

        if line.strip():
            pending_description.append(line.rstrip())
        else:
            pending_description.append("")

    if in_table:
        yield description_for_current, parse_table(table_lines)


def clean_description(lines: list[str]) -> list[str]:
    while lines and not lines[0].strip():
        lines = lines[1:]
    while lines and not lines[-1].strip():
        lines = lines[:-1]
    return lines


def heading_path(node: OrgNode) -> str:
    parts: list[str] = []
    current: OrgNode | None = node
    while current is not None and current.heading:
        parts.append(current.heading)
        current = current.parent
    return " / ".join(reversed(parts))


def collect_tables(root: OrgNode) -> list[ExtractedTable]:
    extracted: list[ExtractedTable] = []
    for node in root[1:]:  # skip synthetic root
        body = node.body or ""
        if not body.strip():
            continue
        for description, rows in extract_from_body(body):
            if not rows or all(not any(cell for cell in row) for row in rows):
                continue
            extracted.append(
                ExtractedTable(
                    heading=node.heading,
                    full_path=heading_path(node),
                    description=clean_description(description),
                    rows=rows,
                )
            )
    return extracted


def make_sheet_name(heading: str, used: set[str]) -> str:
    name = INVALID_SHEET_CHARS_RE.sub(" ", heading)
    name = WHITESPACE_RE.sub(" ", name).strip()
    if not name:
        name = "Sheet"
    name = name[:31]

    candidate = name
    counter = 2
    while candidate in used:
        suffix = f" ({counter})"
        candidate = name[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate)
    return candidate


def column_widths(rows: list[list[str]]) -> list[int]:
    if not rows:
        return []
    col_count = max(len(row) for row in rows)
    widths = [10] * col_count
    for row in rows:
        for idx, cell in enumerate(row):
            longest_token = max((len(part) for part in cell.split("\n")), default=0)
            widths[idx] = max(widths[idx], longest_token)
    # Cap width so wrapped cells stay readable in Excel.
    return [min(max(w + 2, 10), 60) for w in widths]


def write_xlsx(tables: list[ExtractedTable], output: Path) -> None:
    workbook = xlsxwriter.Workbook(str(output), {"strings_to_formulas": False})

    title_fmt = workbook.add_format(
        {"bold": True, "font_size": 14, "font_color": "#1F2937"}
    )
    path_fmt = workbook.add_format(
        {"italic": True, "font_size": 10, "font_color": "#6B7280"}
    )
    desc_fmt = workbook.add_format(
        {
            "italic": True,
            "font_color": "#374151",
            "text_wrap": True,
            "valign": "top",
        }
    )
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#D9E1F2",
            "border": 1,
            "text_wrap": True,
            "valign": "top",
            "align": "left",
        }
    )
    cell_fmt = workbook.add_format(
        {"border": 1, "text_wrap": True, "valign": "top"}
    )

    used_names: set[str] = set()
    for table in tables:
        sheet_name = make_sheet_name(table.heading, used_names)
        sheet = workbook.add_worksheet(sheet_name)
        sheet.set_default_row(18)

        row_idx = 0
        sheet.write(row_idx, 0, table.heading, title_fmt)
        row_idx += 1
        if table.full_path and table.full_path != table.heading:
            sheet.write(row_idx, 0, table.full_path, path_fmt)
            row_idx += 1
        row_idx += 1  # blank line

        if table.description:
            description_text = "\n".join(table.description)
            col_count = max(len(r) for r in table.rows) if table.rows else 1
            last_col = max(col_count - 1, 0)
            line_count = len(table.description)
            note_height = max(30, 15 * line_count)
            if last_col > 0:
                sheet.merge_range(
                    row_idx, 0, row_idx, last_col, description_text, desc_fmt
                )
            else:
                sheet.write(row_idx, 0, description_text, desc_fmt)
            sheet.set_row(row_idx, note_height)
            row_idx += 2

        widths = column_widths(table.rows)
        for col_idx, width in enumerate(widths):
            sheet.set_column(col_idx, col_idx, width)

        if not table.rows:
            continue

        header = table.rows[0]
        for col_idx, value in enumerate(header):
            sheet.write(row_idx, col_idx, value, header_fmt)
        sheet.set_row(row_idx, 28)
        row_idx += 1

        for data_row in table.rows[1:]:
            for col_idx, value in enumerate(data_row):
                sheet.write(row_idx, col_idx, value, cell_fmt)
            row_idx += 1

        sheet.freeze_panes(row_idx - len(table.rows) + 1, 0)

    workbook.close()


def read_preserved_min(xlsx_path: Path) -> dict[str, float]:
    """Read the MIN column from a previously generated Summary sheet, if any.

    Keyed by chapter Code (stable across chapter reorderings and renames of
    everything except the Code itself).
    """
    if not xlsx_path.exists():
        return {}

    try:
        import openpyxl  # local import — only needed in consolidation mode
    except ImportError as e:
        print(
            f"warning: openpyxl not available ({e}); MIN column cannot be "
            "preserved. Install with `uv sync` or `pip install openpyxl`.",
            file=sys.stderr,
        )
        return {}

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        print(
            f"warning: could not read existing xlsx ({e}); MIN column will "
            "start empty.",
            file=sys.stderr,
        )
        return {}

    if "Summary" not in wb.sheetnames:
        return {}

    ws = wb["Summary"]
    code_col = min_col = None
    for cell in ws[1]:
        if cell.value is None:
            continue
        value = str(cell.value).strip()
        if code_col is None and value == "Code":
            code_col = cell.column
        elif min_col is None and (value == "MIN" or value == "MIN (h)"):
            # Match the editable MIN column exactly. The other "MIN @ …"
            # cost columns share the "MIN" prefix but are derived.
            min_col = cell.column
    if code_col is None or min_col is None:
        return {}

    preserved: dict[str, float] = {}
    for row in ws.iter_rows(min_row=2):
        code_cell = row[code_col - 1]
        min_cell = row[min_col - 1]
        if code_cell.value and min_cell.value not in (None, ""):
            try:
                preserved[str(code_cell.value).strip()] = float(min_cell.value)
            except (TypeError, ValueError):
                continue
    return preserved


def rg_multiplier(uur: float) -> float:
    """Total = Dev × (1.70 + 1.30 × UUR).

    1.70 = 1 Dev + 0.30 BA + 0.30 QA + 0.10 PM
    1.30 = (Dev + QA) / Dev = 1 + 0.30
    """
    return 1.70 + uur * 1.30


def write_summary_sheet(
    workbook: xlsxwriter.Workbook,
    chapters: list[ChapterTotal],
    preserved_min: dict[str, float],
    rate: float,
) -> None:
    """Prepend a `Summary` sheet with per-chapter totals and formula-driven costs."""
    sheet = workbook.add_worksheet("Summary")
    sheet.set_default_row(18)

    m30 = rg_multiplier(0.30)
    m40 = rg_multiplier(0.40)
    m50 = rg_multiplier(0.50)

    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#D9E1F2",
            "border": 1,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
        }
    )
    cell_fmt = workbook.add_format({"border": 1, "valign": "top"})
    cell_wrap_fmt = workbook.add_format(
        {"border": 1, "valign": "top", "text_wrap": True}
    )
    money_fmt = workbook.add_format(
        {"border": 1, "valign": "top", "num_format": '"$"#,##0'}
    )
    money_bold = workbook.add_format(
        {
            "border": 1,
            "bold": True,
            "bg_color": "#FCE4D6",
            "num_format": '"$"#,##0',
        }
    )
    int_bold = workbook.add_format(
        {
            "border": 1,
            "bold": True,
            "bg_color": "#FCE4D6",
        }
    )
    text_bold = workbook.add_format(
        {
            "border": 1,
            "bold": True,
            "bg_color": "#FCE4D6",
        }
    )

    headers = [
        "#",
        "Part",
        "Code",
        "Chapter",
        "Dev (h)",
        "MIN (h)",  # user-editable, preserved across runs
        "Δ (Dev − MIN)",
        f"Dev @ ${rate:.0f}/h UUR30",
        f"MIN @ ${rate:.0f}/h UUR30",
        f"Dev @ ${rate:.0f}/h UUR40",
        f"Dev @ ${rate:.0f}/h UUR50",
        "Notes",
    ]
    for col, h in enumerate(headers):
        sheet.write(0, col, h, header_fmt)
    sheet.set_row(0, 32)

    widths = [4, 6, 14, 42, 8, 8, 12, 16, 16, 16, 16, 42]
    for col, w in enumerate(widths):
        sheet.set_column(col, col, w)

    # Data rows
    for i, ch in enumerate(chapters):
        row = i + 1
        dev = ch.total if ch.total is not None else 0
        min_val = preserved_min.get(ch.code)

        sheet.write_number(row, 0, ch.number, cell_fmt)
        sheet.write_string(row, 1, ch.part, cell_fmt)
        sheet.write_string(row, 2, ch.code, cell_fmt)
        sheet.write_string(row, 3, ch.name, cell_wrap_fmt)
        sheet.write_number(row, 4, dev, cell_fmt)
        if min_val is None:
            sheet.write_blank(row, 5, None, cell_fmt)
        else:
            sheet.write_number(row, 5, min_val, cell_fmt)
        # Formulas (Excel is 1-indexed)
        r = row + 1
        sheet.write_formula(row, 6, f'=IF(F{r}="","",E{r}-F{r})', cell_fmt)
        sheet.write_formula(row, 7, f"=E{r}*{m30:.4f}*{rate}", money_fmt)
        sheet.write_formula(
            row, 8, f'=IF(F{r}="","",F{r}*{m30:.4f}*{rate})', money_fmt
        )
        sheet.write_formula(row, 9, f"=E{r}*{m40:.4f}*{rate}", money_fmt)
        sheet.write_formula(row, 10, f"=E{r}*{m50:.4f}*{rate}", money_fmt)
        note = "; ".join(ch.notes) if ch.notes else ""
        sheet.write_string(row, 11, note, cell_wrap_fmt)

    # Grand total row
    n = len(chapters)
    total_row = n + 1
    r = total_row + 1  # 1-indexed for Excel
    last = n + 1  # 1-indexed
    sheet.write_string(total_row, 2, "TOTAL", text_bold)
    sheet.write_string(total_row, 3, f"All {n} chapters", text_bold)
    sheet.write_formula(total_row, 4, f"=SUM(E2:E{last})", int_bold)
    sheet.write_formula(total_row, 5, f"=SUM(F2:F{last})", int_bold)
    sheet.write_formula(total_row, 6, f"=E{r}-F{r}", int_bold)
    sheet.write_formula(total_row, 7, f"=E{r}*{m30:.4f}*{rate}", money_bold)
    sheet.write_formula(total_row, 8, f"=F{r}*{m30:.4f}*{rate}", money_bold)
    sheet.write_formula(total_row, 9, f"=E{r}*{m40:.4f}*{rate}", money_bold)
    sheet.write_formula(total_row, 10, f"=E{r}*{m50:.4f}*{rate}", money_bold)

    sheet.freeze_panes(1, 0)

    # Legend
    legend = workbook.add_worksheet("Legend")
    title_fmt = workbook.add_format({"bold": True, "font_size": 12})
    label_fmt = workbook.add_format({"bold": True})
    wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})

    legend.set_column(0, 0, 22)
    legend.set_column(1, 1, 18)
    legend.set_column(2, 2, 55)

    legend.write("A1", "RG Simple Formula", title_fmt)
    legend.write("A2", "Total = Dev × 1.70 + UUR × 1.30 × Dev = Dev × (1.70 + 1.30 × UUR)")
    legend.write("A4", "Component", label_fmt)
    legend.write("B4", "Value", label_fmt)
    legend.write("C4", "Notes", label_fmt)
    rows = [
        ("Dev", "as-is", "Development"),
        ("BA", "0.30 × Dev", "Business analyst"),
        ("QA", "0.30 × Dev", "Quality assurance"),
        ("PM", "0.10 × Dev", "Project management"),
        ("UUR 30 %", f"× {m30:.4f}", "Extension calibration (RG default)"),
        ("UUR 40 %", f"× {m40:.4f}", "Mid"),
        ("UUR 50 %", f"× {m50:.4f}", "High"),
        ("", "", ""),
        ("Rate", f"${rate:.2f}/h", "RG standard"),
        ("", "", ""),
        ("MIN column", "user-edited", "Preserved across regenerations, keyed by Code."),
        ("Chapter Code", "stable ID", "Change the Code and the MIN entry becomes orphaned."),
    ]
    for i, (a, b, c) in enumerate(rows, start=5):
        legend.write(f"A{i}", a, label_fmt if a and not a.startswith("UUR") else None)
        legend.write(f"B{i}", b)
        legend.write(f"C{i}", c, wrap_fmt)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="org-to-xlsx",
        description="Extract org-mode tables from a .org file into a multi-sheet .xlsx workbook.",
    )
    parser.add_argument("input", type=Path, help="Source .org file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx file (default: <input>.xlsx next to the source)",
    )
    parser.add_argument(
        "--rate",
        type=float,
        default=60.0,
        help="Hourly rate for the Summary sheet cost columns "
        "(only used when `#+CONSOLIDATE_TOTALS:` is present). Default: 60.",
    )
    args = parser.parse_args(argv)

    if not args.input.exists():
        parser.error(f"Input file not found: {args.input}")

    output = args.output or args.input.with_suffix(".xlsx")

    root = orgparse.load(str(args.input))
    tables = collect_tables(root)

    if not tables:
        print(f"No tables found in {args.input}", file=sys.stderr)
        return 1

    # Consolidation mode is triggered only when the source contains
    # `#+CONSOLIDATE_TOTALS: <marker>`. Otherwise the tool is a pure
    # per-table dumper as before.
    marker = read_directive(args.input)
    if marker:
        preserved_min = read_preserved_min(output)
        chapters = parse_totals(args.input, marker)
        write_xlsx_with_summary(
            tables, chapters, preserved_min, output, marker, args.rate
        )
        print(
            f"Wrote {len(tables)} table sheet(s) plus Summary of "
            f"{len(chapters)} chapter(s) to {output}"
        )
        if preserved_min:
            print(
                f"Preserved MIN values for {len(preserved_min)} chapters "
                "from previous run."
            )
    else:
        write_xlsx(tables, output)
        print(f"Wrote {len(tables)} table(s) to {output}")

    return 0


def write_xlsx_with_summary(
    tables: list[ExtractedTable],
    chapters: list[ChapterTotal],
    preserved_min: dict[str, float],
    output: Path,
    marker: str,
    rate: float,
) -> None:
    """Same as `write_xlsx`, but the workbook opens with a Summary sheet first."""
    workbook = xlsxwriter.Workbook(str(output), {"strings_to_formulas": False})

    # Summary + Legend first
    write_summary_sheet(workbook, chapters, preserved_min, rate)

    # Then the regular per-table sheets
    title_fmt = workbook.add_format(
        {"bold": True, "font_size": 14, "font_color": "#1F2937"}
    )
    path_fmt = workbook.add_format(
        {"italic": True, "font_size": 10, "font_color": "#6B7280"}
    )
    desc_fmt = workbook.add_format(
        {
            "italic": True,
            "font_color": "#374151",
            "text_wrap": True,
            "valign": "top",
        }
    )
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#D9E1F2",
            "border": 1,
            "text_wrap": True,
            "valign": "top",
            "align": "left",
        }
    )
    cell_fmt = workbook.add_format(
        {"border": 1, "text_wrap": True, "valign": "top"}
    )

    used_names: set[str] = {"Summary", "Legend"}
    for table in tables:
        sheet_name = make_sheet_name(table.heading, used_names)
        sheet = workbook.add_worksheet(sheet_name)
        sheet.set_default_row(18)

        row_idx = 0
        sheet.write(row_idx, 0, table.heading, title_fmt)
        row_idx += 1
        if table.full_path and table.full_path != table.heading:
            sheet.write(row_idx, 0, table.full_path, path_fmt)
            row_idx += 1
        row_idx += 1

        if table.description:
            description_text = "\n".join(table.description)
            col_count = max(len(r) for r in table.rows) if table.rows else 1
            last_col = max(col_count - 1, 0)
            line_count = len(table.description)
            note_height = max(30, 15 * line_count)
            if last_col > 0:
                sheet.merge_range(
                    row_idx, 0, row_idx, last_col, description_text, desc_fmt
                )
            else:
                sheet.write(row_idx, 0, description_text, desc_fmt)
            sheet.set_row(row_idx, note_height)
            row_idx += 2

        widths = column_widths(table.rows)
        for col_idx, width in enumerate(widths):
            sheet.set_column(col_idx, col_idx, width)

        if not table.rows:
            continue

        header = table.rows[0]
        for col_idx, value in enumerate(header):
            sheet.write(row_idx, col_idx, value, header_fmt)
        sheet.set_row(row_idx, 28)
        row_idx += 1

        for data_row in table.rows[1:]:
            for col_idx, value in enumerate(data_row):
                sheet.write(row_idx, col_idx, value, cell_fmt)
            row_idx += 1

        sheet.freeze_panes(row_idx - len(table.rows) + 1, 0)

    workbook.close()


if __name__ == "__main__":
    raise SystemExit(main())
