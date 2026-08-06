from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Any


@dataclass
class ReconciliationConfig:
    """Controls matching without changing accounting values."""

    cad_currency: str = "CAD"
    usd_currency: str = "USD"
    usd_to_cad_rate: Decimal = Decimal("1")
    floating_tolerance: Decimal = Decimal("0.000001")


@dataclass
class Transaction:
    plant: str
    currency: str
    date: date | None
    journal: str
    batch: str
    references: str
    description: str
    original_amount: Decimal
    converted_amount: Decimal
    status: str = "Unmatched"
    match_id: str = ""
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "Plant": self.plant,
            "Currency": self.currency,
            "Date": self.date,
            "Journal": self.journal,
            "Batch": self.batch,
            "References": self.references,
            "Description": self.description,
            "Original Amount": self.original_amount,
            "Converted Amount": self.converted_amount,
            "Status": self.status,
            "Match ID": self.match_id,
            "Source File": self.source_file,
            "Source Sheet": self.source_sheet,
            "Source Row": self.source_row,
        }
