"""Extract org-mode tables from an .org file into an .xlsx workbook.

Each table becomes its own sheet. The sheet name is derived from the
nearest enclosing heading. Any text that appears between the heading and
the table is written above the table as an italic note block, so the
context for each table travels with it into Excel.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
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


@dataclass
class ExtractedTable:
    heading: str
    full_path: str
    description: list[str]
    rows: list[list[str]]


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
    args = parser.parse_args(argv)

    if not args.input.exists():
        parser.error(f"Input file not found: {args.input}")

    output = args.output or args.input.with_suffix(".xlsx")

    root = orgparse.load(str(args.input))
    tables = collect_tables(root)

    if not tables:
        print(f"No tables found in {args.input}", file=sys.stderr)
        return 1

    write_xlsx(tables, output)
    print(f"Wrote {len(tables)} table(s) to {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
