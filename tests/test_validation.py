from pathlib import Path

from src.anchor_detection import HeaderMatch
from src.sheet_processing import DataBlock, WorksheetResult
from src.validation import validate_table


def _make_worksheet_result(header_values, matched_pair_cols, rows, start_col=1):
    header_match = HeaderMatch(
        row_index=5, start_col=start_col, end_col=start_col + len(header_values) - 1,
        header_values=header_values, matched_pair=("A", "B"),
        matched_pair_cols=matched_pair_cols, occurrence_rows=[5],
    )
    data = DataBlock(rows=rows, row_count=len(rows))
    return WorksheetResult(
        source_file=Path("f.xlsx"), ingestion_type="risk", worksheet_name="Sheet1",
        status="extracted", header_match=header_match, data=data,
    )


def test_valid_table_passes_all_rules():
    result = _make_worksheet_result(
        ["Policy Number", "Inception Date", "Region"],
        {"Policy Number": 1, "Inception Date": 2},
        [["POL-1", "2026-01-01", "UK"], ["POL-2", "2026-02-01", "EU"]],
    )

    validation = validate_table(result)

    assert validation.passed is True
    assert validation.warnings == []
    assert all(r.passed for r in validation.rule_results)


def test_no_data_rows_fails_has_data_rows_rule():
    result = _make_worksheet_result(
        ["Policy Number", "Inception Date"],
        {"Policy Number": 1, "Inception Date": 2},
        [],
    )

    validation = validate_table(result)

    assert validation.passed is False
    outcome = next(r for r in validation.rule_results if r.name == "has_data_rows")
    assert outcome.passed is False


def test_blank_anchor_columns_in_a_row_fails_validation():
    result = _make_worksheet_result(
        ["Policy Number", "Inception Date", "Region"],
        {"Policy Number": 1, "Inception Date": 2},
        [["POL-1", "2026-01-01", "UK"], [None, None, "should not have gotten here"]],
    )

    validation = validate_table(result)

    assert validation.passed is False
    outcome = next(r for r in validation.rule_results if r.name == "no_blank_interior_rows")
    assert outcome.passed is False


def test_inconsistent_row_width_fails_validation():
    result = _make_worksheet_result(
        ["Policy Number", "Inception Date", "Region"],
        {"Policy Number": 1, "Inception Date": 2},
        [["POL-1", "2026-01-01", "UK"], ["POL-2", "2026-02-01"]],  # missing 3rd column
    )

    validation = validate_table(result)

    assert validation.passed is False
    outcome = next(r for r in validation.rule_results if r.name == "consistent_row_width")
    assert outcome.passed is False


def test_duplicate_header_labels_is_warning_only_not_a_failure():
    result = _make_worksheet_result(
        ["Policy Number", "Premium", "Premium"],
        {"Policy Number": 1},
        [["POL-1", 100, 10]],
    )

    validation = validate_table(result)

    assert validation.passed is True
    assert "duplicate header labels present" in validation.warnings


def test_sparse_header_fails_header_has_content():
    result = _make_worksheet_result(
        ["Policy Number", None],
        {"Policy Number": 1},
        [["POL-1", None]],
    )

    validation = validate_table(result)

    outcome = next(r for r in validation.rule_results if r.name == "header_has_content")
    assert outcome.passed is False
