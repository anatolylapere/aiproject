import csv
import re
from pathlib import Path

from src.anchor_detection import HeaderMatch
from src.logging_utils import NullProcessingLogger, ProcessingLogger, build_processing_log_workbook
from src.sheet_processing import DataBlock, WorksheetResult
from src.validation import RuleOutcome, ValidationResult, validate_table


def _read_rows(csv_path):
    with open(csv_path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------------------
# ProcessingLogger: persistent, append-only CSV step log
# ---------------------------------------------------------------------------

def test_start_run_creates_csv_with_header_and_run_start_row(tmp_path):
    csv_path = tmp_path / "logs" / "processing_log.csv"
    logger = ProcessingLogger.start_run(csv_path)

    assert csv_path.exists()
    rows = _read_rows(csv_path)
    assert rows[0]["step"] == "run_start"
    assert rows[0]["run_id"] == logger.run_id


def test_second_run_appends_without_rewriting_header(tmp_path):
    csv_path = tmp_path / "logs" / "processing_log.csv"
    logger_1 = ProcessingLogger.start_run(csv_path)
    logger_2 = ProcessingLogger.start_run(csv_path)

    rows = _read_rows(csv_path)
    assert [r["step"] for r in rows] == ["run_start", "run_start"]
    assert logger_1.run_id != logger_2.run_id

    # exactly one header line in the raw file (DictReader already consumed it once)
    raw_lines = csv_path.read_text(encoding="utf-8").splitlines()
    assert raw_lines[0].startswith("run_id,")
    assert raw_lines.count(raw_lines[0]) == 1


def test_log_step_appends_a_row_with_expected_fields(tmp_path):
    logger = ProcessingLogger.start_run(tmp_path / "processing_log.csv")

    logger.log_step(Path("f.xlsx"), "risk", "Risk_Main", "anchor_detection", "ok", "matched pair X/Y")

    rows = _read_rows(logger.csv_path)
    step_row = next(r for r in rows if r["step"] == "anchor_detection")
    assert step_row["source_file"] == str(Path("f.xlsx"))
    assert step_row["ingestion_type"] == "risk"
    assert step_row["worksheet_name"] == "Risk_Main"
    assert step_row["status"] == "ok"
    assert step_row["detail"] == "matched pair X/Y"
    assert re.fullmatch(r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}", step_row["timestamp"])


def test_log_run_end_appends_row(tmp_path):
    logger = ProcessingLogger.start_run(tmp_path / "processing_log.csv")
    logger.log_run_end()

    rows = _read_rows(logger.csv_path)
    assert rows[-1]["step"] == "run_end"


def test_password_check_step_detail_only_carries_boolean_flag(tmp_path):
    # Never log the actual password value - only whether one was required.
    logger = ProcessingLogger.start_run(tmp_path / "processing_log.csv")

    logger.log_step(Path("f.xlsx"), "risk", None, "password_check", "ok", "password_protected=True")

    rows = _read_rows(logger.csv_path)
    step_row = next(r for r in rows if r["step"] == "password_check")
    assert "SuperSecretPassword123" not in step_row["detail"]
    assert step_row["detail"] == "password_protected=True"


def test_null_processing_logger_is_a_no_op():
    NullProcessingLogger().log_step(Path("f.xlsx"), "risk", "Sheet1", "anchor_detection", "ok", "detail")
    # no assertion needed beyond "does not raise" - confirms process_worksheet etc. can be
    # called directly in unit tests without a logger and produce no side effects.


def test_validate_table_with_default_logger_does_not_raise():
    header_match = HeaderMatch(
        row_index=1, start_col=1, end_col=2, header_values=["A", "B"],
        matched_pair=("A", "B"), matched_pair_cols={"A": 1, "B": 2}, occurrence_rows=[1],
    )
    result = WorksheetResult(
        source_file=Path("f.xlsx"), ingestion_type="risk", worksheet_name="Sheet1",
        status="extracted", header_match=header_match, data=DataBlock(rows=[["1", "2"]], row_count=1),
    )
    validate_table(result)  # default NULL_LOGGER - must not raise or require a logger


# ---------------------------------------------------------------------------
# build_processing_log_workbook: per-output-folder processing_log.xlsx
# ---------------------------------------------------------------------------

def _extracted_result(
    source_file, worksheet_name, header_values, matched_pair_cols, row_count, passed=True, contract_code_year="",
):
    header_match = HeaderMatch(
        row_index=1, start_col=1, end_col=len(header_values), header_values=header_values,
        matched_pair=("A", "B"), matched_pair_cols=matched_pair_cols, occurrence_rows=[1],
    )
    data = DataBlock(rows=[["v"] * len(header_values)] * row_count, row_count=row_count)
    rule_results = [RuleOutcome("has_data_rows", passed, "")]
    validation = ValidationResult(passed=passed, rule_results=rule_results, warnings=[])
    return WorksheetResult(
        source_file=source_file, ingestion_type="risk", worksheet_name=worksheet_name,
        status="extracted", header_match=header_match, data=data, validation=validation,
        output_path=Path(f"output/risk/{source_file}"), contract_code_year=contract_code_year,
    )


def _skipped_result(source_file, worksheet_name, status):
    return WorksheetResult(
        source_file=source_file, ingestion_type="risk", worksheet_name=worksheet_name, status=status,
    )


def test_granular_sheet_includes_every_worksheet_regardless_of_outcome():
    results = [
        _extracted_result("f1.xlsx", "Main", ["Policy Number", "Premium"], {"A": 1, "B": 2}, 2),
        _skipped_result("f1.xlsx", "Notes", "skipped_no_header"),
        _skipped_result("f1.xlsx", "NoData", "skipped_no_data"),
    ]

    workbook = build_processing_log_workbook([{"source_file": "f1.xlsx", "opened_ok": True, "password_required": False}], results)

    ws = workbook["Granular"]
    sheet_names = [row[1].value for row in ws.iter_rows(min_row=2)]
    assert sheet_names == ["Main", "Notes", "NoData"]
    statuses = [row[2].value for row in ws.iter_rows(min_row=2)]
    assert statuses == ["extracted", "skipped_no_header", "skipped_no_data"]


def test_granular_sheet_records_contract_code_year_for_traceability():
    results = [
        _extracted_result(
            "f1.xlsx", "Main", ["Policy Number", "Premium"], {"A": 1, "B": 2}, 2, contract_code_year="BW0197221A",
        ),
        _skipped_result("f1.xlsx", "Notes", "skipped_no_header"),
    ]

    workbook = build_processing_log_workbook([{"source_file": "f1.xlsx", "opened_ok": True, "password_required": False}], results)

    ws = workbook["Granular"]
    header = [c.value for c in ws[1]]
    assert "Contract Code Year" in header
    contract_col = header.index("Contract Code Year")
    values = [row[contract_col].value for row in ws.iter_rows(min_row=2)]
    assert values == ["BW0197221A", ""]  # skipped worksheet never ran detection - blank


def test_overview_sheet_counts_files_and_worksheet_statuses():
    results = [
        _extracted_result("f1.xlsx", "Main", ["A", "B"], {"A": 1, "B": 2}, 2, passed=True),
        _extracted_result("f1.xlsx", "Secondary", ["A", "B"], {"A": 1, "B": 2}, 1, passed=False),
        _skipped_result("f1.xlsx", "Notes", "skipped_no_header"),
    ]
    file_summaries = [
        {"source_file": "f1.xlsx", "opened_ok": True, "password_required": True},
        {"source_file": "f2.xlsx", "opened_ok": False, "password_required": False},
    ]

    workbook = build_processing_log_workbook(file_summaries, results)

    ws = workbook["Overview"]
    metrics = {row[0].value: row[1].value for row in ws.iter_rows(min_row=2)}
    assert metrics["Total files processed"] == 2
    assert metrics["Files opened successfully"] == 1
    assert metrics["Files that failed to open"] == 1
    assert metrics["Files with password protection"] == 1
    assert metrics["Worksheets extracted"] == 2
    assert metrics["Worksheets skipped (no header)"] == 1
    assert metrics["Tables passed validation"] == 1
    assert metrics["Tables failed validation"] == 1


def test_column_mapping_sheet_counts_distinct_file_sheet_occurrences_and_leaves_sink_blank():
    # "Premium" appears in both f1/Main and f1/Secondary -> frequency 2.
    # "Region" appears only once, and twice within the SAME header (duplicate) -> still frequency 1.
    results = [
        _extracted_result("f1.xlsx", "Main", ["Policy Number", "Premium", "Region", "Region"], {"A": 1, "B": 2}, 1),
        _extracted_result("f1.xlsx", "Secondary", ["Claim Number", "Premium"], {"A": 1, "B": 2}, 1),
        _skipped_result("f1.xlsx", "Notes", "skipped_no_header"),  # contributes nothing
    ]

    workbook = build_processing_log_workbook([{"source_file": "f1.xlsx", "opened_ok": True, "password_required": False}], results)

    ws = workbook["ColumnMapping"]
    assert [c.value for c in ws[1]] == ["Source", "Frequency", "Sink"]
    rows = {row[0].value: (row[1].value, row[2].value) for row in ws.iter_rows(min_row=2)}
    assert rows["Premium"] == (2, None)
    assert rows["Region"] == (1, None)
    assert rows["Policy Number"] == (1, None)
    assert rows["Claim Number"] == (1, None)
