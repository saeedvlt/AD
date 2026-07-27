"""Unpivot base-detail rows from the 2025 Margin Analysis workbook.

Designed for the Budget Database Toolkit.  It processes Windsor, Cambridge,
Montreal, and Ithaca and returns a normalized pandas DataFrame.
"""

from __future__ import annotations

from datetime import date, datetime
from numbers import Number
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


LOCATIONS = ("Windsor", "Cambridge", "Montreal", "Ithaca")
SECTION_NAMES = {
    "sales": "Sales",
    "material": "Material",
    "labour": "Labour",
    "labor": "Labour",
    "overhead": "Overhead",
    "gross margin": "Gross Margin",
}


def clean_text(value: Any) -> str:
    """Return safely trimmed text for a worksheet value."""
    return "" if value is None else str(value).strip()


def is_number(value: Any) -> bool:
    """True for spreadsheet numeric values, excluding booleans."""
    return isinstance(value, Number) and not isinstance(value, bool)


def find_month_columns(ws: Worksheet) -> dict[int, datetime]:
    """Find the first contiguous 12-month date series in the top of a sheet."""
    for row in range(1, min(ws.max_row, 15) + 1):
        for start_column in range(1, ws.max_column - 10):
            values = [ws.cell(row, start_column + offset).value for offset in range(12)]
            if all(isinstance(value, (datetime, date)) for value in values):
                periods = [
                    datetime(value.year, value.month, 1)
                    if isinstance(value, date) and not isinstance(value, datetime)
                    else datetime(value.year, value.month, 1)
                    for value in values
                ]
                if [period.month for period in periods] == list(range(1, 13)):
                    return {
                        start_column + offset: period
                        for offset, period in enumerate(periods)
                    }
    raise ValueError(f"Could not find a Jan–Dec date series on worksheet '{ws.title}'.")


def is_derived_line(label: str) -> bool:
    """Exclude calculated totals, ratios, and reconciliation lines."""
    normalized = label.casefold()
    return (
        normalized.startswith("total")
        or normalized.startswith("per ")
        or normalized.startswith("sales per")
        or normalized == "diff"
        or normalized.startswith("adjust to actual")
        or normalized.startswith("gross margin")
        or normalized.endswith("%")
        or " per gp" in normalized
        or "financial statement" in normalized
        or normalized in {"f/s", "fabs only", "plug"}
    )


def has_monthly_amount(ws: Worksheet, row: int, month_columns: dict[int, datetime]) -> bool:
    """Keep a base row only when at least one of its 12 months has a value."""
    return any(
        is_number(ws.cell(row, column).value) and ws.cell(row, column).value != 0
        for column in month_columns
    )


def convert(uploaded_file: Any) -> pd.DataFrame:
    """Return base-detail margin analysis data in long format.

    Derived report rows (totals, ratios, gross margin, differences, and
    ``Adjust to Actual``) are intentionally excluded to prevent double-counting.
    Manual adjustment lines, such as ``Adjust: Rebates``, are retained.
    """
    workbook = load_workbook(uploaded_file, data_only=True)
    missing_sheets = [name for name in LOCATIONS if name not in workbook.sheetnames]
    if missing_sheets:
        raise ValueError("Missing required worksheet(s): " + ", ".join(missing_sheets))

    records: list[dict[str, Any]] = []

    for location in LOCATIONS:
        ws = workbook[location]
        month_columns = find_month_columns(ws)
        current_section: str | None = None

        for row in range(1, ws.max_row + 1):
            label = clean_text(ws.cell(row, 1).value)
            normalized_label = label.casefold()

            if normalized_label in SECTION_NAMES:
                current_section = SECTION_NAMES[normalized_label]
                continue

            # Gross Margin and its following rows are report calculations,
            # rather than base Sales/Material/Labour/Overhead detail.
            if current_section == "Gross Margin":
                continue

            if not label or current_section is None or is_derived_line(label):
                continue
            if not has_monthly_amount(ws, row, month_columns):
                continue

            for column, period in month_columns.items():
                value = ws.cell(row, column).value
                amount = round(float(value), 2) if is_number(value) else 0.0
                records.append(
                    {
                        "Location": location,
                        "Section": current_section,
                        "Line Item": label,
                        "Period": period,
                        "Month": period.strftime("%m - %b"),
                        "Amount": amount,
                    }
                )

    columns = ["Location", "Section", "Line Item", "Period", "Month", "Amount"]
    return pd.DataFrame(records, columns=columns)
