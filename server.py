"""Minimal MCP tools for local Excel workbooks."""

from contextlib import closing
import os
from pathlib import Path
from typing import Any

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from openpyxl import load_workbook

mcp = FastMCP("Excel")


def workbook_path(path: str) -> Path:
    """Optionally confine workbook access to the folder shared by the HTTP launcher."""
    shared_root = os.environ.get("KB_EXCEL_ROOT")
    if not shared_root:
        return Path(path)
    root = Path(shared_root).resolve()
    candidate = Path(path)
    if not candidate.is_absolute():
        candidate = root / candidate
    candidate = candidate.resolve()
    if not candidate.is_relative_to(root):
        raise ToolError("Workbook path is outside the shared Excel folder.")
    if candidate.suffix.lower() != ".xlsx":
        raise ToolError("Only .xlsx files can be read from the shared Excel folder.")
    if not candidate.is_file():
        raise ToolError("Workbook not found in the shared Excel folder.")
    return candidate


@mcp.tool
def list_excel_sheets(path: str) -> list[str]:
    """List worksheets in an XLSX file. In shared mode, path is relative to the shared folder."""
    with closing(load_workbook(workbook_path(path), read_only=True, data_only=True)) as workbook:
        return workbook.sheetnames


@mcp.tool
def read_excel_rows(
    path: str, sheet: str, start_row: int = 1, limit: int = 100
) -> list[dict[str, Any]]:
    """Read up to limit rows (1–1000), starting at a 1-based Excel row number.

    Returns row numbers and cell values, including headers and blank cells.
    Formulas return their cached values; formulas are not calculated.
    In shared mode, path is relative to the shared Excel folder.
    """
    if start_row < 1 or not 1 <= limit <= 1000:
        raise ValueError("start_row must be >= 1 and limit must be between 1 and 1000")
    with closing(load_workbook(workbook_path(path), read_only=True, data_only=True)) as workbook:
        worksheet = workbook[sheet]
        if worksheet.max_row is not None and start_row > worksheet.max_row:
            return []
        end_row = start_row + limit - 1
        if worksheet.max_row is not None:
            end_row = min(end_row, worksheet.max_row)
        return [
            {"row": number, "values": list(values)}
            for number, values in enumerate(
                worksheet.iter_rows(
                    min_row=start_row, max_row=end_row, values_only=True
                ),
                start=start_row,
            )
        ]


@mcp.tool
def search_excel(
    path: str, query: str, sheet: str | None = None, limit: int = 100
) -> list[dict[str, Any]]:
    """Find cells containing query (case-insensitive) in one or all worksheets.

    Returns up to limit matches (1–1000), with sheet, cell address and value.
    Searches cached formula values, without calculating formulas.
    In shared mode, path is relative to the shared Excel folder.
    """
    if not query or not 1 <= limit <= 1000:
        raise ValueError("query must not be empty and limit must be between 1 and 1000")
    matches = []
    needle = query.casefold()
    with closing(load_workbook(workbook_path(path), read_only=True, data_only=True)) as workbook:
        worksheets = [workbook[sheet]] if sheet is not None else workbook.worksheets
        for worksheet in worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value is not None and needle in str(cell.value).casefold():
                        matches.append(
                            {"sheet": worksheet.title, "cell": cell.coordinate, "value": cell.value}
                        )
                        if len(matches) >= limit:
                            return matches
    return matches


if __name__ == "__main__":
    mcp.run()
