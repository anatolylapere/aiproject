"""Validation rules for extracted tables.

Operates only on already-extracted WorksheetResult data (header_match/data) - no
worksheet/file I/O - so it is testable with plain constructed objects.
"""

from dataclasses import dataclass, field

from src.logging_utils import NULL_LOGGER


@dataclass
class RuleOutcome:
    name: str
    passed: bool
    detail: str = ""


@dataclass
class ValidationResult:
    passed: bool
    rule_results: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def validate_table(worksheet_result, logger=NULL_LOGGER):
    header_match = worksheet_result.header_match
    data = worksheet_result.data

    rule_results = []
    warnings = []

    non_empty_header_cells = sum(1 for v in header_match.header_values if v is not None)
    rule_results.append(RuleOutcome(
        "header_has_content",
        non_empty_header_cells >= 2,
        f"{non_empty_header_cells} non-empty header cells",
    ))

    rule_results.append(RuleOutcome(
        "has_data_rows",
        data.row_count >= 1,
        f"row_count={data.row_count}",
    ))

    anchor_cols = sorted(set(header_match.matched_pair_cols.values()))
    blank_anchor_rows = [
        row for row in data.rows
        if all(row[c - header_match.start_col] is None for c in anchor_cols)
    ]
    rule_results.append(RuleOutcome(
        "no_blank_interior_rows",
        len(blank_anchor_rows) == 0,
        f"{len(blank_anchor_rows)} rows with blank anchor columns",
    ))

    expected_width = header_match.end_col - header_match.start_col + 1
    bad_width_rows = [row for row in data.rows if len(row) != expected_width]
    rule_results.append(RuleOutcome(
        "consistent_row_width",
        len(bad_width_rows) == 0,
        f"{len(bad_width_rows)} rows with unexpected width",
    ))

    non_null_labels = [v for v in header_match.header_values if v is not None]
    if len(non_null_labels) != len(set(non_null_labels)):
        warnings.append("duplicate header labels present")

    overall_passed = all(r.passed for r in rule_results)

    detail = "all rules passed" if overall_passed else "failed: " + ", ".join(
        r.name for r in rule_results if not r.passed
    )
    logger.log_step(
        worksheet_result.source_file, worksheet_result.ingestion_type, worksheet_result.worksheet_name,
        "validation", "pass" if overall_passed else "fail", detail,
    )

    return ValidationResult(passed=overall_passed, rule_results=rule_results, warnings=warnings)
