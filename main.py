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
from src.contract_detection import load_contract_codes
from src.logging_utils import ProcessingLogger, build_processing_log_workbook
from src.sheet_processing import (
    is_risk_filename, process_worksheet, worksheet_name_has_date, write_combined_workbook,
)
from src.validation import validate_table

PROJECT_ROOT = Path(__file__).resolve().parent
FILELIST_PATH = PROJECT_ROOT / "input" / "filelist.csv"
ANCHOR_PAIRS_PATH = PROJECT_ROOT / "config" / "anchor_pairs.json"
CONTRACT_CODES_PATH = PROJECT_ROOT / "config" / "contract_codes.json"
PASSWORD_CONFIG_PATH = PROJECT_ROOT / "config" / "password.json"
OUTPUT_ROOT = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"


def main():
    entries = file_access.read_filelist(FILELIST_PATH)
    anchors = load_anchor_pairs(ANCHOR_PAIRS_PATH)
    contract_config = load_contract_codes(CONTRACT_CODES_PATH)
    logger = ProcessingLogger.start_run(LOGS_DIR / "processing_log.csv")

    worksheet_results_by_type = {"risk": [], "claims": []}
    file_summaries_by_type = {"risk": [], "claims": []}

    for entry in entries:
        worksheet_results, file_summary = process_entry(entry, anchors, contract_config, logger)
        worksheet_results_by_type[entry.ingestion_type].extend(worksheet_results)
        file_summaries_by_type[entry.ingestion_type].append(file_summary)

    for ingestion_type in ("risk", "claims"):
        write_processing_log_workbook(
            ingestion_type, file_summaries_by_type[ingestion_type], worksheet_results_by_type[ingestion_type],
        )

    logger.log_run_end()


def process_entry(entry, anchors, contract_config, logger):
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
        result
        for worksheet in workbook.worksheets
        for result in _process_and_validate(worksheet, entry, anchors, contract_config, logger)
    ]

    write_outputs(entry, worksheet_results, logger)

    return worksheet_results, file_summary


def _process_and_validate(worksheet, entry, anchors, contract_config, logger):
    # a worksheet normally yields one result, but the Property Sec/Province row-by-row
    # fallback can split it into several (see sheet_processing.process_worksheet)
    results = process_worksheet(
        worksheet, worksheet.title, entry.source_file, entry.ingestion_type,
        anchors[entry.ingestion_type], contract_config, logger=logger,
    )
    for result in results:
        if result.status == "extracted":
            result.validation = validate_table(result, logger=logger)
    return results


def _group_code(results):
    """The single contract code shared by every resolved result in this group, or None
    if none resolved or they disagree - a group with no resolved code, or with
    disagreeing codes, gets no code-based folder/suffix rather than guessing which
    contract it belongs to.
    """
    codes = {r.contract_code_year for r in results if r.contract_code_year}
    return codes.pop() if len(codes) == 1 else None


def _write_grouped(base_dir, filename_stem, results, logger):
    """Write results as one workbook. If every result agrees on a contract code, the
    file is nested under base_dir/{code}/ and the code is appended to the filename;
    otherwise it's written flat at base_dir/{filename_stem}.xlsx.
    """
    code = _group_code(results)
    if code:
        write_combined_workbook(base_dir / code / f"{filename_stem}__{code}.xlsx", results, logger=logger)
    else:
        write_combined_workbook(base_dir / f"{filename_stem}.xlsx", results, logger=logger)


def _write_sectioned_groups(base_dir, stem, results, logger):
    """Write every result with a resolved section to its own file, grouped by exact
    ContractCodeYear (so worksheets sharing a section share a file). Returns the
    remaining (non-sectioned) results for the caller to combine/split further.
    """
    sectioned = [r for r in results if r.contract_section is not None]
    groups = {}
    for result in sectioned:
        groups.setdefault(result.contract_code_year, []).append(result)
    for group in groups.values():
        _write_grouped(base_dir, stem, group, logger)
    return [r for r in results if r not in sectioned]


def write_outputs(entry, worksheet_results, logger):
    """Write extracted+validated tables to output/{risk|premium|claims}/ (asymmetric
    grouping per the approved plan). Both Risk (split into output/risk/ or
    output/premium/ by a filename keyword check) and Claims split out section-bearing
    worksheets (grouped by ContractCodeYear, so worksheets sharing a section share a
    file) into their own files first; Claims additionally splits date-named worksheets
    into their own files. Whatever's left combines into one file per source file. A
    worksheet that is both section-bearing and date-named is claimed by the section
    group first. Any output whose sheets agree on a single contract code is nested
    under a {code}/ folder and has that code appended to its filename.
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
        base_dir = OUTPUT_ROOT / ("risk" if is_risk_filename(stem) else "premium")
        remaining = _write_sectioned_groups(base_dir, stem, passed, logger)
        if remaining:
            _write_grouped(base_dir, stem, remaining, logger)
        return

    remaining = _write_sectioned_groups(OUTPUT_ROOT / "claims", stem, passed, logger)
    dated = [r for r in remaining if worksheet_name_has_date(r.worksheet_name)]
    no_date = [r for r in remaining if r not in dated]
    for result in dated:
        _write_grouped(OUTPUT_ROOT / "claims", f"{stem}__{result.worksheet_name}", [result], logger)
    if no_date:
        _write_grouped(OUTPUT_ROOT / "claims", stem, no_date, logger)


def write_processing_log_workbook(ingestion_type, file_summaries, worksheet_results):
    if not file_summaries:
        return  # no files of this type were processed this run
    output_dir = OUTPUT_ROOT / ingestion_type
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = build_processing_log_workbook(file_summaries, worksheet_results)
    workbook.save(output_dir / "processing_log.xlsx")


if __name__ == "__main__":
    main()
