import csv
from pathlib import Path

import openpyxl

import main
from src.anchor_detection import load_anchor_pairs
from src.file_access import FileListEntry
from src.logging_utils import ProcessingLogger

TEST_FILES_DIR = Path(__file__).resolve().parent / "test files"
CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "anchor_pairs.json"


def _run(monkeypatch, tmp_path, entries):
    output_root = tmp_path / "output"
    logs_dir = tmp_path / "logs"
    monkeypatch.setattr(main, "OUTPUT_ROOT", output_root)
    monkeypatch.setattr(main, "LOGS_DIR", logs_dir)

    anchors = load_anchor_pairs(CONFIG_PATH)
    logger = ProcessingLogger.start_run(logs_dir / "processing_log.csv")

    worksheet_results_by_type = {"risk": [], "claims": []}
    file_summaries_by_type = {"risk": [], "claims": []}
    for entry in entries:
        worksheet_results, file_summary = main.process_entry(entry, anchors, logger)
        worksheet_results_by_type[entry.ingestion_type].extend(worksheet_results)
        file_summaries_by_type[entry.ingestion_type].append(file_summary)

    for ingestion_type in ("risk", "claims"):
        main.write_processing_log_workbook(
            ingestion_type, file_summaries_by_type[ingestion_type], worksheet_results_by_type[ingestion_type],
        )
    logger.log_run_end()

    return output_root, logger.csv_path, worksheet_results_by_type


def _read_csv_rows(csv_path):
    with open(csv_path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _statuses(worksheet_results):
    return {r.worksheet_name: r.status for r in worksheet_results}


def test_risk_fixture_combines_into_one_workbook(tmp_path, monkeypatch):
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file1_Risk_test.xlsx", ingestion_type="risk")

    output_root, csv_path, results_by_type = _run(monkeypatch, tmp_path, [entry])

    output_file = output_root / "risk" / "file1_Risk_test.xlsx"
    assert output_file.exists()
    workbook = openpyxl.load_workbook(output_file)
    assert set(workbook.sheetnames) == {"Risk_Main", "Risk_AltAnchor"}

    assert _statuses(results_by_type["risk"]) == {
        "Risk_Main": "extracted",
        "Risk_AltAnchor": "extracted",
        "Risk_NoData": "skipped_no_data",
        "Notes": "skipped_no_header",
    }

    # step-level CSV log carries granular evidence, not just a final status
    rows = _read_csv_rows(csv_path)
    anchor_steps = [r for r in rows if r["step"] == "anchor_detection" and r["worksheet_name"] == "Risk_Main"]
    assert anchor_steps[0]["status"] == "ok"
    assert "Policy Number/Inception Date" in anchor_steps[0]["detail"]


def test_claims_fixture_combines_since_no_sheet_name_has_a_date(tmp_path, monkeypatch):
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file2_Claims_test.xlsx", ingestion_type="claims")

    output_root, _, results_by_type = _run(monkeypatch, tmp_path, [entry])

    combined_file = output_root / "claims" / "file2_Claims_test.xlsx"
    assert combined_file.exists()
    workbook = openpyxl.load_workbook(combined_file)
    assert set(workbook.sheetnames) == {"Claims_Main", "Claims_Secondary"}

    claims_dir = output_root / "claims"
    assert sorted(p.name for p in claims_dir.iterdir()) == sorted(["file2_Claims_test.xlsx", "processing_log.xlsx"])

    assert _statuses(results_by_type["claims"]) == {
        "Claims_Main": "extracted",
        "Claims_Secondary": "extracted",
        "Claims_NoData": "skipped_no_data",
        "ReadMe": "skipped_no_header",
    }


def test_claims_month_name_date_in_sheet_name_gets_its_own_file(tmp_path, monkeypatch):
    # file3_Claims_test.xlsx: same fixture as file2, except the main sheet is named
    # "Claims_Main Aug 2026" - a month-name-style date, not the numeric ISO format.
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file3_Claims_test.xlsx", ingestion_type="claims")

    output_root, _, results_by_type = _run(monkeypatch, tmp_path, [entry])

    claims_dir = output_root / "claims"
    own_file = claims_dir / "file3_Claims_test__Claims_Main Aug 2026.xlsx"
    combined_file = claims_dir / "file3_Claims_test.xlsx"

    assert own_file.exists()
    assert combined_file.exists()

    own_workbook = openpyxl.load_workbook(own_file)
    assert own_workbook.sheetnames == ["Claims_Main Aug 2026"]
    own_ws = own_workbook["Claims_Main Aug 2026"]
    assert [c.value for c in own_ws[1]] == [
        "Claim Number", "Policy Number Date", "Policy Number", "Claim Amount", "Claim Amount_1", "Status", "metadata",
    ]
    assert own_ws.max_row == 5  # header + 4 data rows (CLM-001, 002, 003, 999)

    combined_workbook = openpyxl.load_workbook(combined_file)
    assert combined_workbook.sheetnames == ["Claims_Secondary"]

    assert _statuses(results_by_type["claims"]) == {
        "Claims_Main Aug 2026": "extracted",
        "Claims_Secondary": "extracted",
        "Claims_NoData": "skipped_no_data",
        "ReadMe": "skipped_no_header",
    }


def test_rerun_overwrites_deterministically_and_never_touches_source(tmp_path, monkeypatch):
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file1_Risk_test.xlsx", ingestion_type="risk")
    source_bytes_before = entry.source_file.read_bytes()

    output_root_1, _, _ = _run(monkeypatch, tmp_path, [entry])
    output_file = output_root_1 / "risk" / "file1_Risk_test.xlsx"

    output_root_2, _, _ = _run(monkeypatch, tmp_path, [entry])

    # same path reused (overwrite, not versioned), and source file untouched
    assert output_root_1 == output_root_2
    assert entry.source_file.read_bytes() == source_bytes_before

    workbook = openpyxl.load_workbook(output_file)
    assert set(workbook.sheetnames) == {"Risk_Main", "Risk_AltAnchor"}


def test_claims_worksheet_with_date_in_name_gets_its_own_file(tmp_path, monkeypatch):
    # Synthetic fixture: one sheet named with a date, one without - exercises the
    # split-by-date-in-name grouping rule not present in the real sample workbooks.
    source_path = tmp_path / "input" / "Claims_Synthetic.xlsx"
    source_path.parent.mkdir(parents=True)
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    dated_sheet = workbook.create_sheet("Claims_2026-07")
    dated_sheet.append(["Source", "Test"])
    dated_sheet.append(["Claim Number", "Claim Amount"])
    dated_sheet.append(["CLM-1", 100])
    dated_sheet.append(["CLM-2", 200])

    no_date_sheet = workbook.create_sheet("Claims_NoDate")
    no_date_sheet.append(["Source", "Test"])
    no_date_sheet.append(["Claim Number", "Claim Amount"])
    no_date_sheet.append(["CLM-3", 300])

    workbook.save(source_path)

    entry = FileListEntry(source_file=source_path, ingestion_type="claims")
    output_root, _, _ = _run(monkeypatch, tmp_path, [entry])

    claims_dir = output_root / "claims"
    own_file = claims_dir / "Claims_Synthetic__Claims_2026-07.xlsx"
    combined_file = claims_dir / "Claims_Synthetic.xlsx"

    assert own_file.exists()
    assert combined_file.exists()

    own_workbook = openpyxl.load_workbook(own_file)
    assert own_workbook.sheetnames == ["Claims_2026-07"]

    combined_workbook = openpyxl.load_workbook(combined_file)
    assert combined_workbook.sheetnames == ["Claims_NoDate"]


def test_processing_log_xlsx_written_to_each_output_folder(tmp_path, monkeypatch):
    risk_entry = FileListEntry(source_file=TEST_FILES_DIR / "file1_Risk_test.xlsx", ingestion_type="risk")
    claims_entry = FileListEntry(source_file=TEST_FILES_DIR / "file2_Claims_test.xlsx", ingestion_type="claims")

    output_root, _, _ = _run(monkeypatch, tmp_path, [risk_entry, claims_entry])

    risk_log = openpyxl.load_workbook(output_root / "risk" / "processing_log.xlsx")
    claims_log = openpyxl.load_workbook(output_root / "claims" / "processing_log.xlsx")

    assert risk_log.sheetnames == ["Granular", "Overview", "ColumnMapping"]
    assert claims_log.sheetnames == ["Granular", "Overview", "ColumnMapping"]

    # risk's log must not contain claims worksheets and vice versa
    risk_sheet_names = [row[1].value for row in risk_log["Granular"].iter_rows(min_row=2)]
    assert set(risk_sheet_names) == {"Risk_Main", "Risk_AltAnchor", "Risk_NoData", "Notes"}
    claims_sheet_names = [row[1].value for row in claims_log["Granular"].iter_rows(min_row=2)]
    assert set(claims_sheet_names) == {"Claims_Main", "Claims_Secondary", "Claims_NoData", "ReadMe"}


def test_processing_log_never_written_for_a_type_with_no_files_processed(tmp_path, monkeypatch):
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file1_Risk_test.xlsx", ingestion_type="risk")

    output_root, _, _ = _run(monkeypatch, tmp_path, [entry])

    assert not (output_root / "claims").exists()


def test_processing_log_csv_persists_and_accumulates_across_runs(tmp_path, monkeypatch):
    entry = FileListEntry(source_file=TEST_FILES_DIR / "file1_Risk_test.xlsx", ingestion_type="risk")

    _, csv_path_1, _ = _run(monkeypatch, tmp_path, [entry])
    _, csv_path_2, _ = _run(monkeypatch, tmp_path, [entry])

    assert csv_path_1 == csv_path_2
    rows = _read_csv_rows(csv_path_1)
    run_ids = {r["run_id"] for r in rows if r["step"] == "run_start"}
    assert len(run_ids) == 2  # both runs' evidence accumulated in the same file
