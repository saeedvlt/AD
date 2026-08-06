from __future__ import annotations

import pandas as pd

from .models import Transaction


STANDARD_COLUMNS = ["Plant", "Currency", "Date", "Journal", "Batch", "References", "Description", "Original Amount", "Converted Amount", "Status", "Match ID", "Source File", "Source Sheet", "Source Row"]


def transactions_frame(transactions: list[Transaction]) -> pd.DataFrame:
    frame = pd.DataFrame([t.to_dict() for t in transactions], columns=STANDARD_COLUMNS)
    if not frame.empty:
        frame["Date"] = pd.to_datetime(frame["Date"], errors="coerce").dt.date
    return frame


def matches_frame(matches: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(matches, columns=["Match ID", "Left Amount", "Right Amount", "Difference", "Left Description", "Right Description"])
