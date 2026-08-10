"""Header row detection against configured Risk/Claims anchor pairs."""

import json
from dataclasses import dataclass


@dataclass
class HeaderMatch:
    row_index: int
    start_col: int
    end_col: int
    header_values: list
    matched_pair: tuple
    matched_pair_cols: dict
    occurrence_rows: list


def load_anchor_pairs(config_path):
    """Load and validate config/anchor_pairs.json. Returns {"risk": [(a, b), ...], "claims": [...]}."""
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError(f"Anchor pairs config must be a JSON object: {config_path}")

    result = {}
    for ingestion_type in ("risk", "claims"):
        pairs = data.get(ingestion_type)
        if not isinstance(pairs, list):
            raise ValueError(f"Anchor pairs config missing '{ingestion_type}' list: {config_path}")
        normalized_pairs = []
        for pair in pairs:
            if not (isinstance(pair, list) and len(pair) == 2):
                raise ValueError(f"Anchor pair must be a 2-element list: {pair!r}")
            normalized_pairs.append((pair[0], pair[1]))
        result[ingestion_type] = normalized_pairs
    return result


def _normalize(value):
    if value is None:
        return None
    return str(value).strip().lower()


def _row_values(worksheet, row_index, max_col):
    return [worksheet.cell(row=row_index, column=col).value for col in range(1, max_col + 1)]


def _row_matches_pair(row_values, pair):
    normalized_values = {_normalize(v) for v in row_values if v is not None}
    return _normalize(pair[0]) in normalized_values and _normalize(pair[1]) in normalized_values


def _row_extent(row_values):
    """Return 1-based (start_col, end_col) spanning the non-empty cells of a row, or None if blank."""
    non_empty_cols = [i + 1 for i, v in enumerate(row_values) if v is not None]
    if not non_empty_cols:
        return None
    return non_empty_cols[0], non_empty_cols[-1]


def _pair_cols(row_values, pair):
    """Map each pair string to the (first) 1-based column it was found in."""
    targets = {_normalize(pair[0]): pair[0], _normalize(pair[1]): pair[1]}
    cols = {}
    for i, value in enumerate(row_values):
        normalized = _normalize(value)
        if normalized in targets:
            label = targets[normalized]
            if label not in cols:
                cols[label] = i + 1
    return cols


def find_header_row(worksheet, anchor_pairs):
    """Pair-priority, early-stop, all-occurrences-collected header detection.

    Tries each pair in config order; the first pair that matches any row wins (later
    pairs are never tried). Every row matched by the winning pair is collected as an
    occurrence, top to bottom - the first occurrence is the canonical header row that
    defines the column framework (start_col/end_col/header_values) for extraction.

    Returns a HeaderMatch, or None if no pair matches anywhere in the sheet.
    """
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    all_row_values = [_row_values(worksheet, r, max_col) for r in range(1, max_row + 1)]

    for pair in anchor_pairs:
        occurrence_rows = [
            row_index
            for row_index, row_values in enumerate(all_row_values, start=1)
            if _row_matches_pair(row_values, pair)
        ]
        if not occurrence_rows:
            continue

        first_row_values = all_row_values[occurrence_rows[0] - 1]
        extent = _row_extent(first_row_values)
        start_col, end_col = extent
        header_values = first_row_values[start_col - 1:end_col]
        matched_pair_cols = _pair_cols(first_row_values, pair)

        return HeaderMatch(
            row_index=occurrence_rows[0],
            start_col=start_col,
            end_col=end_col,
            header_values=header_values,
            matched_pair=pair,
            matched_pair_cols=matched_pair_cols,
            occurrence_rows=occurrence_rows,
        )

    return None
