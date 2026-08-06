from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from .models import ReconciliationConfig, Transaction


ALIASES = {
    "plant": ["plant", "location", "entity", "company"],
    "currency": ["currency", "curr", "iso currency"],
    "date": ["date", "transaction date", "trans date", "posting date"],
    "journal": ["journal", "journal id", "journal name", "journal entry"],
    "batch": ["batch", "batch id", "batch number", "batch no"],
    "references": ["reference", "references", "ref", "invoice", "invoice number", "invoice no"],
    "description": ["description", "memo", "detail", "details", "transaction description"],
    "amount": ["amount", "total", "debit", "credit", "net amount", "transaction amount"],
}


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _norm_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", _clean(value).lower()).strip()


def _parse_decimal(value: object) -> Decimal | None:
    text = _clean(value)
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("() ").replace(",", "").replace("$", "").replace("€", "")
    if text.endswith("-"):
        text = "-" + text[:-1]
    try:
        result = Decimal(text)
        return -result if negative else result
    except InvalidOperation:
        return None


def _parse_date(value: object) -> date | None:
    if value is None or _clean(value) == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def _find_header(df: pd.DataFrame) -> int:
    best_row, best_score = 0, -1
    for row_index in range(min(len(df), 30)):
        cells = {_norm_header(v) for v in df.iloc[row_index].tolist()}
        score = sum(any(alias == cell or alias in cell for alias in aliases) for aliases in ALIASES.values() for cell in cells)
        if score > best_score:
            best_row, best_score = row_index, score
    return best_row


def _column_map(columns: list[object]) -> dict[str, object]:
    result: dict[str, object] = {}
    for column in columns:
        normalized = _norm_header(column)
        for key, aliases in ALIASES.items():
            if key not in result and any(alias == normalized or alias in normalized for alias in aliases):
                result[key] = column
    return result


def _read_excel(source: str | Path | bytes | BinaryIO) -> tuple[str, dict[str, pd.DataFrame]]:
    if isinstance(source, (str, Path)):
        label = Path(source).name
        sheets = pd.read_excel(source, sheet_name=None, header=None, dtype=object)
    else:
        label = getattr(source, "name", "uploaded_workbook.xlsx")
        raw = source if isinstance(source, bytes) else source.read()
        sheets = pd.read_excel(BytesIO(raw), sheet_name=None, header=None, dtype=object)
    return label, sheets


def load_transactions(source: str | Path | bytes | BinaryIO, currency: str, config: ReconciliationConfig, plant: str = "") -> list[Transaction]:
    """Load all ledger-like sheets and normalize rows into the shared schema."""
    source_file, sheets = _read_excel(source)
    transactions: list[Transaction] = []
    for sheet_name, raw in sheets.items():
        if raw.dropna(how="all").empty:
            continue
        header_row = _find_header(raw)
        header = raw.iloc[header_row].tolist()
        data = raw.iloc[header_row + 1:].copy()
        data.columns = header
        mapping = _column_map(header)
        if "amount" not in mapping:
            continue
        for source_row, (_, row) in enumerate(data.iterrows(), start=header_row + 2):
            amount = _parse_decimal(row.get(mapping["amount"]))
            description = _clean(row.get(mapping.get("description", "")))
            if amount is None or not description and all(_clean(v) == "" for v in row.tolist()):
                continue
            if description.lower() in {"beginning balance", "subtotal", "total", "ending balance"}:
                continue
            row_currency = _clean(row.get(mapping.get("currency", ""))) or currency
            factor = config.usd_to_cad_rate if row_currency.upper() == config.usd_currency.upper() else Decimal("1")
            transactions.append(Transaction(
                plant=_clean(row.get(mapping.get("plant", ""))) or plant,
                currency=row_currency.upper(),
                date=_parse_date(row.get(mapping.get("date", ""))),
                journal=_clean(row.get(mapping.get("journal", ""))),
                batch=_clean(row.get(mapping.get("batch", ""))),
                references=_clean(row.get(mapping.get("references", ""))),
                description=description,
                original_amount=amount,
                converted_amount=amount * factor,
                source_file=source_file,
                source_sheet=str(sheet_name),
                source_row=source_row,
            ))
    return transactions
