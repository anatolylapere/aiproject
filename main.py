"""Workflow orchestrator for the Claims & Risk Excel ingestion pipeline.

Reads input/filelist.csv, processes each source workbook's worksheets, validates
extracted tables, writes clean output, and records processing evidence: a persistent
step-level CSV log (logs/processing_log.csv, appended to on every run) plus a
processing_log.xlsx rebuilt fresh each run inside each output/{risk|claims}/ folder.
Never modifies source files. See CLAUDE.md for the full requirements.
"""

from pathlib import Path

from src import file_access
from src.anchor_detection import load_anchor_pairs
from src.logging_utils import ProcessingLogger, build_processing_log_workbook
from src.sheet_processing import process_worksheet, worksheet_name_has_date, write_combined_workbook
from src.validation import validate_table

PROJECT_ROOT = Path(__file__).resolve().parent
FILELIST_PATH = PROJECT_ROOT / "input" / "filelist.csv"
ANCHOR_PAIRS_PATH = PROJECT_ROOT / "config" / "anchor_pairs.json"
PASSWORD_CONFIG_PATH = PROJECT_ROOT / "config" / "password.json"
OUTPUT_ROOT = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"


def main():
    entries = file_access.read_filelist(FILELIST_PATH)
    anchors = load_anchor_pairs(ANCHOR_PAIRS_PATH)
    logger = ProcessingLogger.start_run(LOGS_DIR / "processing_log.csv")

    worksheet_results_by_type = {"risk": [], "claims": []}
    file_summaries_by_type = {"risk": [], "claims": []}

    for entry in entries:
        worksheet_results, file_summary = process_entry(entry, anchors, logger)
        worksheet_results_by_type[entry.ingestion_type].extend(worksheet_results)
        file_summaries_by_type[entry.ingestion_type].append(file_summary)

    for ingestion_type in ("risk", "claims"):
        write_processing_log_workbook(
            ingestion_type, file_summaries_by_type[ingestion_type], worksheet_results_by_type[ingestion_type],
        )

    logger.log_run_end()


def process_entry(entry, anchors, logger):
    file_summary = {
        "source_file": entry.source_file, "opened_ok": False, "password_required": False,
    }

    try:
        workbook, password_required = file_access.open_source_file(entry, PASSWORD_CONFIG_PATH, logger)
    except file_access.WorkbookOpenError:
        return [], file_summary  # failure already logged by open_source_file

    file_summary["opened_ok"] = True
    file_summary["password_required"] = password_required

    worksheet_results = [
        _process_and_validate(worksheet, entry, anchors, logger)
        for worksheet in workbook.worksheets
    ]

    write_outputs(entry, worksheet_results, logger)

    return worksheet_results, file_summary


def _process_and_validate(worksheet, entry, anchors, logger):
    result = process_worksheet(
        worksheet, worksheet.title, entry.source_file, entry.ingestion_type,
        anchors[entry.ingestion_type], logger=logger,
    )
    if result.status == "extracted":
        result.validation = validate_table(result, logger=logger)
    return result


def write_outputs(entry, worksheet_results, logger):
    """Write extracted+validated tables to output/{risk|claims}/ (asymmetric grouping
    per the approved plan - Risk always combines per source file; Claims combines
    no-date-named worksheets the same way but gives date-named ones their own file).
    """
    passed = [r for r in worksheet_results if r.status == "extracted" and r.validation.passed]
    failed = [r for r in worksheet_results if r.status == "extracted" and not r.validation.passed]
    for result in failed:
        logger.log_step(result.source_file, result.ingestion_type, result.worksheet_name,
                         "output_write", "skip", "suppressed: failed validation")

    if not passed:
        return

    stem = entry.source_file.stem
    if entry.ingestion_type == "risk":
        write_combined_workbook(OUTPUT_ROOT / "risk" / f"{stem}.xlsx", passed, logger=logger)
        return

    dated = [r for r in passed if worksheet_name_has_date(r.worksheet_name)]
    no_date = [r for r in passed if r not in dated]
    for result in dated:
        write_combined_workbook(
            OUTPUT_ROOT / "claims" / f"{stem}__{result.worksheet_name}.xlsx", [result], logger=logger,
        )
    if no_date:
        write_combined_workbook(OUTPUT_ROOT / "claims" / f"{stem}.xlsx", no_date, logger=logger)


def write_processing_log_workbook(ingestion_type, file_summaries, worksheet_results):
    if not file_summaries:
        return  # no files of this type were processed this run
    output_dir = OUTPUT_ROOT / ingestion_type
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = build_processing_log_workbook(file_summaries, worksheet_results)
    workbook.save(output_dir / "processing_log.xlsx")


if __name__ == "__main__":
    main()
