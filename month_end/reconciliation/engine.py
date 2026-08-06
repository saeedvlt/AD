from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from itertools import combinations
from math import comb
from uuid import uuid4

from .models import ReconciliationConfig, Transaction


def _new_match_id() -> str:
    return f"M-{uuid4().hex[:10].upper()}"


def _sum_abs(transactions: list[Transaction] | tuple[Transaction, ...]) -> Decimal:
    return sum((abs(transaction.converted_amount) for transaction in transactions), Decimal("0"))


def _same_side(transactions: list[Transaction] | tuple[Transaction, ...]) -> bool:
    nonzero = [transaction.converted_amount for transaction in transactions if transaction.converted_amount != 0]
    return bool(nonzero) and (all(value > 0 for value in nonzero) or all(value < 0 for value in nonzero))


def _opposite_sides(left: list[Transaction] | tuple[Transaction, ...], right: list[Transaction] | tuple[Transaction, ...]) -> bool:
    left_values = [transaction.converted_amount for transaction in left if transaction.converted_amount != 0]
    right_values = [transaction.converted_amount for transaction in right if transaction.converted_amount != 0]
    return bool(left_values and right_values) and _same_side(left) and _same_side(right) and left_values[0] * right_values[0] < 0


def _record_match(match_type: str, left: list[Transaction] | tuple[Transaction, ...], right: list[Transaction] | tuple[Transaction, ...]) -> dict:
    match_id = _new_match_id()
    for transaction in (*left, *right):
        transaction.status = "Matched"
        transaction.match_id = match_id
    left_amount = sum((transaction.converted_amount for transaction in left), Decimal("0"))
    right_amount = sum((transaction.converted_amount for transaction in right), Decimal("0"))
    return {
        "Match ID": match_id,
        "Match Type": match_type,
        "Left Count": len(left),
        "Right Count": len(right),
        "Left Amount": left_amount,
        "Right Amount": right_amount,
        "Difference": left_amount + right_amount,
        "Left Description": " | ".join(transaction.description for transaction in left if transaction.description),
        "Right Description": " | ".join(transaction.description for transaction in right if transaction.description),
    }


def exact_one_to_one_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    """Match exact converted amounts with opposite accounting signs."""
    right_by_amount: dict[Decimal, list[int]] = defaultdict(list)
    for index, transaction in enumerate(right):
        right_by_amount[abs(transaction.converted_amount)].append(index)
    used_right: set[int] = set()
    matches: list[dict] = []
    for left_txn in left:
        candidate_index = None
        for amount, indexes in right_by_amount.items():
            if abs(abs(left_txn.converted_amount) - amount) <= config.floating_tolerance:
                candidate_index = next(
                    (
                        index for index in indexes
                        if index not in used_right
                        and _opposite_sides([left_txn], [right[index]])
                    ),
                    None,
                )
                if candidate_index is not None:
                    break
        if candidate_index is None:
            continue
        used_right.add(candidate_index)
        matches.append(_record_match("1:1", [left_txn], [right[candidate_index]]))
    return [transaction for transaction in left if transaction.status != "Matched"], [transaction for index, transaction in enumerate(right) if index not in used_right], matches


def _find_group(target: list[Transaction] | tuple[Transaction, ...], candidates: list[Transaction], target_is_left: bool, config: ReconciliationConfig, min_group_size: int) -> tuple[Transaction, ...] | None:
    target_amount = _sum_abs(target)
    target_sign = next((transaction.converted_amount for transaction in target if transaction.converted_amount != 0), Decimal("0"))
    if target_sign == 0:
        return None
    for group_size in range(min_group_size, config.max_group_size + 1):
        if comb(len(candidates), group_size) > config.max_group_combinations:
            continue
        for candidate_group in combinations(candidates, group_size):
            if not _same_side(candidate_group):
                continue
            candidate_sign = next(transaction.converted_amount for transaction in candidate_group if transaction.converted_amount != 0)
            if target_sign * candidate_sign >= 0:
                continue
            if abs(target_amount - _sum_abs(candidate_group)) <= config.floating_tolerance:
                return candidate_group
    return None


def keyword_group_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    """Match all same-file transactions containing a configured keyword to one opposite-side entry."""
    matches: list[dict] = []
    remaining_left = list(left)
    remaining_right = list(right)
    keywords = [keyword.strip().lower() for keyword in config.group_keywords if keyword.strip()]
    for keyword in keywords:
        groups: dict[tuple[str, str], list[Transaction]] = defaultdict(list)
        for transaction in remaining_left:
            searchable = f"{transaction.description} {transaction.references}".lower()
            if keyword in searchable:
                groups[(transaction.source_file, transaction.source_sheet)].append(transaction)
        for group in groups.values():
            if not 2 <= len(group) <= config.max_group_size or not _same_side(group):
                continue
            candidate = next(
                (
                    transaction for transaction in remaining_right
                    if _opposite_sides(group, [transaction])
                    and abs(_sum_abs(group) - abs(transaction.converted_amount)) <= config.floating_tolerance
                ),
                None,
            )
            if candidate is None:
                continue
            matches.append(_record_match(f"keyword {keyword} many:1", group, [candidate]))
            remaining_left = [transaction for transaction in remaining_left if transaction not in group]
            remaining_right.remove(candidate)
    return remaining_left, remaining_right, matches


def exact_one_to_many_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    matches: list[dict] = []
    remaining_right = list(right)
    for left_txn in left:
        group = _find_group([left_txn], remaining_right, True, config, 2)
        if group is not None:
            matches.append(_record_match("1:many", [left_txn], list(group)))
            remaining_right = [transaction for transaction in remaining_right if transaction not in group]
    remaining_left = [transaction for transaction in left if transaction.status != "Matched"]
    return remaining_left, remaining_right, matches


def exact_many_to_one_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    matches: list[dict] = []
    remaining_left = list(left)
    for right_txn in right:
        group = _find_group([right_txn], remaining_left, False, config, 2)
        if group is not None:
            matches.append(_record_match("many:1", list(group), [right_txn]))
            remaining_left = [transaction for transaction in remaining_left if transaction not in group]
    remaining_right = [transaction for transaction in right if transaction.status != "Matched"]
    return remaining_left, remaining_right, matches


def exact_many_to_many_match(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> tuple[list[Transaction], list[Transaction], list[dict]]:
    matches: list[dict] = []
    remaining_left = list(left)
    remaining_right = list(right)
    for left_size in range(2, config.max_group_size + 1):
        if comb(len(remaining_left), left_size) > config.max_group_combinations:
            continue
        changed = True
        while changed:
            changed = False
            for left_group in combinations(remaining_left, left_size):
                right_group = _find_group(left_group, remaining_right, True, config, 2)
                if right_group is None:
                    continue
                matches.append(_record_match("many:many", list(left_group), list(right_group)))
                remaining_left = [transaction for transaction in remaining_left if transaction not in left_group]
                remaining_right = [transaction for transaction in remaining_right if transaction not in right_group]
                changed = True
                break
    return remaining_left, remaining_right, matches


def near_match_suggestions(left: list[Transaction], right: list[Transaction], config: ReconciliationConfig) -> list[dict]:
    """Suggest close 1:1 pairs without marking them as reconciled."""
    suggestions: list[dict] = []
    for left_txn in left:
        candidates = [
            right_txn for right_txn in right
            if _opposite_sides([left_txn], [right_txn])
            and abs(abs(left_txn.converted_amount) - abs(right_txn.converted_amount)) <= config.near_match_threshold
        ]
        for right_txn in sorted(candidates, key=lambda transaction: abs(abs(left_txn.converted_amount) - abs(transaction.converted_amount))):
            residual = left_txn.converted_amount + right_txn.converted_amount
            left_txn.status = "Near Match"
            right_txn.status = "Near Match"
            suggestions.append({
                "Match ID": f"N-{uuid4().hex[:10].upper()}",
                "Match Type": "Near 1:1",
                "Left Count": 1,
                "Right Count": 1,
                "Left Amount": left_txn.converted_amount,
                "Right Amount": right_txn.converted_amount,
                "Difference": residual,
                "Left Description": left_txn.description,
                "Right Description": right_txn.description,
            })
    return suggestions


def reconcile(cad: list[Transaction], usd: list[Transaction], config: ReconciliationConfig) -> dict:
    unmatched_cad, unmatched_usd, matches = exact_one_to_one_match(cad, usd, config)
    unmatched_cad, unmatched_usd, grouped = keyword_group_match(unmatched_cad, unmatched_usd, config)
    matches.extend(grouped)
    unmatched_cad, unmatched_usd, grouped = exact_one_to_many_match(unmatched_cad, unmatched_usd, config)
    matches.extend(grouped)
    unmatched_cad, unmatched_usd, grouped = exact_many_to_one_match(unmatched_cad, unmatched_usd, config)
    matches.extend(grouped)
    unmatched_cad, unmatched_usd, grouped = exact_many_to_many_match(unmatched_cad, unmatched_usd, config)
    matches.extend(grouped)
    near_matches = near_match_suggestions(unmatched_cad, unmatched_usd, config)
    return {
        "transactions": cad + usd,
        "matches": matches,
        "near_matches": near_matches,
        "unmatched_cad": unmatched_cad,
        "unmatched_usd": unmatched_usd,
    }
