from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from uuid import uuid4

from .models import ReconciliationConfig, Transaction


def _amount_key(value: Decimal) -> Decimal:
    return value


def exact_one_to_one_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    """Match exact converted amounts, using tolerance only for floating arithmetic."""
    right_by_amount: dict[Decimal, list[int]] = defaultdict(list)
    for index, transaction in enumerate(right):
        right_by_amount[_amount_key(transaction.converted_amount)].append(index)
    used_right: set[int] = set()
    matches: list[dict] = []
    for left_txn in left:
        candidate_index = None
        for amount, indexes in right_by_amount.items():
            if abs(left_txn.converted_amount - amount) <= config.floating_tolerance:
                candidate_index = next((i for i in indexes if i not in used_right), None)
                if candidate_index is not None:
                    break
        if candidate_index is None:
            continue
        right_txn = right[candidate_index]
        used_right.add(candidate_index)
        match_id = f"M-{uuid4().hex[:10].upper()}"
        left_txn.status = right_txn.status = "Matched"
        left_txn.match_id = right_txn.match_id = match_id
        matches.append({
            "Match ID": match_id,
            "Left Amount": left_txn.converted_amount,
            "Right Amount": right_txn.converted_amount,
            "Difference": left_txn.converted_amount - right_txn.converted_amount,
            "Left Description": left_txn.description,
            "Right Description": right_txn.description,
        })
    unmatched_right = [txn for index, txn in enumerate(right) if index not in used_right]
    unmatched_left = [txn for txn in left if txn.status != "Matched"]
    return unmatched_left, unmatched_right, matches


def reconcile(cad: list[Transaction], usd: list[Transaction], config: ReconciliationConfig) -> dict:
    unmatched_cad, unmatched_usd, matches = exact_one_to_one_match(cad, usd, config)
    return {
        "transactions": cad + usd,
        "matches": matches,
        "unmatched_cad": unmatched_cad,
        "unmatched_usd": unmatched_usd,
    }
