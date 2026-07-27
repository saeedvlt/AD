"""Unpivot base-detail Margin Analysis data and add requested percentages."""

from __future__ import annotations

from datetime import date, datetime
from numbers import Number
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


LOCATIONS = ("Windsor", "Cambridge", "Montreal", "Ithaca")
MARKETS = {
    "Windsor": "Canada", "Cambridge": "Canada", "Montreal": "Canada",
    "Ithaca": "United States", "Madison": "United States",
}
SECTION_NAMES = {
    "sales": "Sales", "material": "Material", "labour": "Labour",
    "labor": "Labour", "overhead": "Overhead", "gross margin": "Gross Margin",
}


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def is_number(value: Any) -> bool:
    return isinstance(value, Number) and not isinstance(value, bool)


def find_month_columns(ws: Worksheet) -> dict[int, datetime]:
    """Locate the first contiguous Jan–Dec date header block."""
    for row in range(1, min(ws.max_row, 15) + 1):
        for start in range(1, ws.max_column - 10):
            headers = [ws.cell(row, start + offset).value for offset in range(12)]
            if not all(isinstance(value, (date, datetime)) for value in headers):
                continue
            periods = [datetime(value.year, value.month, 1) for value in headers]
            if [period.month for period in periods] == list(range(1, 13)):
                return {start + offset: period for offset, period in enumerate(periods)}
    raise ValueError(f"Could not find the Jan–Dec date headers on '{ws.title}'.")


def has_monthly_amount(ws: Worksheet, row: int, columns: dict[int, datetime]) -> bool:
    return any(
        is_number(ws.cell(row, column).value) and ws.cell(row, column).value != 0
        for column in columns
    )


def is_derived_line(label: str) -> bool:
    """Exclude totals, ratios, reconciliations, and Gross Margin calculations."""
    text = label.casefold()
    return (
        text.startswith("total") or text.startswith("per ")
        or text.startswith("sales per") or text == "diff"
        or text.startswith("adjust to actual") or text.startswith("gross margin")
        or text.endswith("%") or " per gp" in text or "financial statement" in text
        or text in {"f/s", "fabs only", "plug"}
    )


def percent(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    """Return percentages on a 0–100 scale; divide-by-zero is blank."""
    return (numerator.div(denominator.where(denominator != 0)) * 100).round(2)


def add_percentages(data: pd.DataFrame, location_sales: pd.DataFrame) -> pd.DataFrame:
    """Add market share, location-sales, and matching-Sales cost percentages."""
    output = data.merge(location_sales, on=["Location", "Period"], how="left")

    # 1: Location / same line item across the location's Canada or US market.
    market_total = (
        output.groupby(["Market", "Section", "Line Item", "Period"], as_index=False)["Amount"]
        .sum().rename(columns={"Amount": "_market_line_total"})
    )
    output = output.merge(
        market_total, on=["Market", "Section", "Line Item", "Period"], how="left"
    )
    output["Location Share of Market Line Item %"] = percent(
        output["Amount"], output["_market_line_total"]
    )

    # 2: Every line / total Sales of that exact location and month.
    output["Line Item % of Location Sales"] = percent(
        output["Amount"], output["Location Total Sales"]
    )

    # 3–5: Material, Labour, and Overhead / Sales of the matching line item.
    by_section = (
        output.groupby(["Location", "Line Item", "Period", "Section"], as_index=False)["Amount"]
        .sum().pivot_table(
            index=["Location", "Line Item", "Period"],
            columns="Section", values="Amount", aggfunc="sum",
        ).reset_index()
    )
    by_section.columns.name = None
    for section in ("Sales", "Material", "Labour", "Overhead"):
        if section not in by_section:
            by_section[section] = pd.NA
    by_section = by_section.rename(columns={
        "Sales": "_sales", "Material": "_material",
        "Labour": "_labour", "Overhead": "_overhead",
    })
    output = output.merge(by_section, on=["Location", "Line Item", "Period"], how="left")

    output["Material % of Matching Sales"] = pd.NA
    output["Labour % of Matching Sales"] = pd.NA
    output["Overhead % of Matching Sales"] = pd.NA
    for section, amount_column, result_column in (
        ("Material", "_material", "Material % of Matching Sales"),
        ("Labour", "_labour", "Labour % of Matching Sales"),
        ("Overhead", "_overhead", "Overhead % of Matching Sales"),
    ):
        mask = output["Section"].eq(section)
        output.loc[mask, result_column] = percent(
            output.loc[mask, amount_column], output.loc[mask, "_sales"]
        )

    return output.drop(columns=["_market_line_total", "_sales", "_material", "_labour", "_overhead"])


def convert(uploaded_file: Any) -> pd.DataFrame:
    """Unpivot four sites; exclude derived rows; add requested percentage columns."""
    workbook = load_workbook(uploaded_file, data_only=True)
    missing = [name for name in LOCATIONS if name not in workbook.sheetnames]
    if missing:
        raise ValueError("Missing required worksheet(s): " + ", ".join(missing))

    records: list[dict[str, Any]] = []
    sales_totals: list[dict[str, Any]] = []

    for location in LOCATIONS:
        ws = workbook[location]
        months = find_month_columns(ws)
        section: str | None = None

        for row in range(1, ws.max_row + 1):
            label = clean_text(ws.cell(row, 1).value)
            normalized = label.casefold()

            if normalized in SECTION_NAMES:
                section = SECTION_NAMES[normalized]
                continue

            # Retain Sales Total only as a calculation denominator.
            if section == "Sales" and normalized == "total":
                for column, period in months.items():
                    value = ws.cell(row, column).value
                    if is_number(value):
                        sales_totals.append({
                            "Location": location,
                            "Period": period,
                            "Location Total Sales": round(float(value), 2),
                        })
                continue

            if section == "Gross Margin" or not label or section is None:
                continue
            if is_derived_line(label) or not has_monthly_amount(ws, row, months):
                continue

            for column, period in months.items():
                value = ws.cell(row, column).value
                records.append({
                    "Location": location,
                    "Market": MARKETS[location],
                    "Section": section,
                    "Line Item": label,
                    "Period": period,
                    "Month": period.strftime("%m - %b"),
                    "Amount": round(float(value), 2) if is_number(value) else 0.0,
                })

    columns = ["Location", "Market", "Section", "Line Item", "Period", "Month", "Amount"]
    data = pd.DataFrame(records, columns=columns)
    sales = pd.DataFrame(sales_totals, columns=["Location", "Period", "Location Total Sales"])
    return add_percentages(data, sales) if not data.empty else data
