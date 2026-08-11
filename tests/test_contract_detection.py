from pathlib import Path

import pytest

from src.contract_detection import ContractCodeConfig, detect_contract_code_year, load_contract_codes
from src.metadata_extraction import MetadataBlock

CONTRACT_CODES_PATH = Path(__file__).resolve().parent.parent / "config" / "contract_codes.json"

# Load once and reuse everywhere below, so these tests always exercise whatever is
# actually configured in contract_codes.json - no hand-copied config to drift out of
# sync with it. Current shape: BW01972 (aliases incl. "Hardy"/"CNA", sections A/B) and
# BW01973 (aliases incl. "HDI", sections A/B).
REAL_CONFIG = load_contract_codes(CONTRACT_CODES_PATH)


def _detect(
    metadata_values=(), header_values=(), data_rows=(), worksheet_name="Sheet1", source_file="file.xlsx",
    metadata_cells=(),
):
    header_values = list(header_values)
    data_rows = [list(r) for r in data_rows]
    return detect_contract_code_year(
        source_file=Path(source_file), worksheet_name=worksheet_name,
        metadata=MetadataBlock(values=list(metadata_values), cells=list(metadata_cells)),
        header_values=header_values, data_rows=data_rows,
        header_row=5, start_col=1, data_row_numbers=list(range(6, 6 + len(data_rows))),
        contract_config=REAL_CONFIG,
    )


# ---------------------------------------------------------------------------
# load_contract_codes
# ---------------------------------------------------------------------------

def test_load_contract_codes_reads_real_config():
    assert "BW01972" in REAL_CONFIG.contracts and "BW01973" in REAL_CONFIG.contracts
    assert "Hardy" in REAL_CONFIG.contracts["BW01972"]["aliases"]
    assert "A" in REAL_CONFIG.contracts["BW01972"]["sections"]
    assert "A" in REAL_CONFIG.contracts["BW01973"]["sections"]


def test_load_contract_codes_rejects_missing_contracts_key(tmp_path):
    path = tmp_path / "contract_codes.json"
    path.write_text('{"foo": {}}', encoding="utf-8")

    with pytest.raises(ValueError):
        load_contract_codes(path)


def test_load_contract_codes_rejects_contract_without_aliases(tmp_path):
    path = tmp_path / "contract_codes.json"
    path.write_text('{"contracts": {"X": {"name": "Y"}}}', encoding="utf-8")

    with pytest.raises(ValueError):
        load_contract_codes(path)


def test_load_contract_codes_rejects_section_without_aliases(tmp_path):
    path = tmp_path / "contract_codes.json"
    path.write_text('{"contracts": {"X": {"aliases": ["X"], "sections": {"A": {}}}}}', encoding="utf-8")

    with pytest.raises(ValueError):
        load_contract_codes(path)


# ---------------------------------------------------------------------------
# detect_contract_code_year: single-tier matches
# ---------------------------------------------------------------------------

def test_metadata_tier_match():
    result = _detect(metadata_values=["Broker", "Hardy Ltd"])

    assert result.contract_code_year == "BW01972"
    assert result.status == "resolved"
    assert result.tier == "metadata"


def test_table_column_tier_match():
    result = _detect(header_values=["Claim Number", "HDI Premium"])

    assert result.contract_code_year == "BW01973"
    assert result.tier == "table_columns_and_values"


def test_table_data_value_tier_match():
    result = _detect(header_values=["Claim Number", "Status"], data_rows=[["CLM-1", "HDI"]])

    assert result.contract_code_year == "BW01973"
    assert result.tier == "table_columns_and_values"


def test_worksheet_name_tier_match():
    result = _detect(worksheet_name="CNA Claims")

    assert result.contract_code_year == "BW01972"
    assert result.tier == "worksheet_name"


def test_filename_tier_match():
    result = _detect(source_file="Hardy_Risk_test.xlsx")

    assert result.contract_code_year == "BW01972"
    assert result.tier == "filename"


# ---------------------------------------------------------------------------
# tier priority: metadata > columns/values > worksheet name > filename
# ---------------------------------------------------------------------------

def test_metadata_tier_wins_over_conflicting_filename():
    # filename says Hardy (BW01972), metadata says HDI (BW01973) - metadata wins,
    # and the filename evidence is never even consulted.
    result = _detect(metadata_values=["HDI"], source_file="Hardy_Risk_test.xlsx")

    assert result.contract_code_year == "BW01973"
    assert result.tier == "metadata"


def test_columns_tier_only_consulted_when_metadata_has_no_match():
    result = _detect(metadata_values=["Broker", "Acme"], header_values=["HDI Premium"])

    assert result.contract_code_year == "BW01973"
    assert result.tier == "table_columns_and_values"


# ---------------------------------------------------------------------------
# code-style alias matching: contains-match (no boundary) since the code itself is
# specific enough on its own - mirrors file_Claims_test.xlsx's "broker ref" value,
# where the contract code is embedded directly in a longer reference string with no
# separators. Word aliases (no digits, e.g. "Hardy"/"CNA") still require isolation.
# ---------------------------------------------------------------------------

def test_code_style_alias_matches_even_when_embedded_in_a_longer_token():
    result = _detect(metadata_values=["broker ref", "b1262bw0197219"])

    assert result.contract_code_year == "BW01972"
    assert result.tier == "metadata"


def test_word_alias_still_requires_isolation_when_embedded_in_a_longer_token():
    # "Hardy" glued directly onto other letters (no separator) must NOT match -
    # word aliases keep the boundary protection the code-style ones no longer have.
    result = _detect(metadata_values=["xHardyx"])

    assert result.status == "not_found"


def test_word_alias_as_isolated_token_within_longer_string_does_match():
    # contrast with the case above: filename tokens ARE isolated by "_" and must match.
    result = _detect(source_file="Hardy_Risk_test.xlsx")

    assert result.contract_code_year == "BW01972"


def test_code_that_does_not_appear_at_all_still_does_not_match():
    # specificity still matters: a code with genuinely different digits must not match
    # just because it shares a prefix with the real one.
    result = _detect(metadata_values=["BW01975"])

    assert result.status == "not_found"


# ---------------------------------------------------------------------------
# section resolution
# ---------------------------------------------------------------------------

def test_section_resolved_via_worksheet_name():
    result = _detect(metadata_values=["Hardy"], worksheet_name="Claims Sec B 2026")

    assert result.contract_code_year == "BW01972B"
    assert result.status == "resolved"


def test_hdi_now_resolves_sections_too():
    # both real contracts have sections configured now
    result = _detect(metadata_values=["HDI"], worksheet_name="Section A")

    assert result.contract_code_year == "BW01973A"


def test_section_not_applicable_when_contract_has_no_sections():
    # synthetic config: a contract with no "sections" key at all must ignore
    # section-shaped evidence entirely, rather than erroring or guessing
    config = ContractCodeConfig(contracts={"X01": {"aliases": ["XCorp"]}})
    result = detect_contract_code_year(
        source_file=Path("file.xlsx"), worksheet_name="Section A",
        metadata=MetadataBlock(values=["XCorp"]), header_values=[], data_rows=[],
        header_row=5, start_col=1, data_row_numbers=[],
        contract_config=config,
    )

    assert result.contract_code_year == "X01"
    assert result.section is None


def test_no_section_evidence_anywhere_yields_base_code_only():
    result = _detect(metadata_values=["Hardy"], worksheet_name="Risk_Main", source_file="Hardy_Risk_test.xlsx")

    assert result.contract_code_year == "BW01972"


def test_section_ambiguous_within_a_tier_drops_section_but_keeps_base():
    result = _detect(metadata_values=["Hardy"], worksheet_name="Section A and Section B")

    assert result.contract_code_year == "BW01972"
    assert result.status == "resolved"
    assert "ambiguous" in result.detail


# ---------------------------------------------------------------------------
# "Property Sec"/"Property Section" column: the cell value itself IS the section code
# (mirrors file1_Risk_ HDI Rolling Total.xlsx's Property Sec column)
# ---------------------------------------------------------------------------

def test_property_sec_column_value_resolves_section():
    result = _detect(
        metadata_values=["Hardy"], header_values=["Policy Number", "Property Sec"],
        data_rows=[["POL-1", "A"]],
    )

    assert result.contract_code_year == "BW01972A"
    assert result.section == "A"


def test_property_section_full_word_header_also_matches():
    result = _detect(
        metadata_values=["Hardy"], header_values=["Policy Number", "Property Section"],
        data_rows=[["POL-1", "B"]],
    )

    assert result.contract_code_year == "BW01972B"


def test_property_sec_column_value_not_a_configured_section_is_ignored():
    result = _detect(
        metadata_values=["Hardy"], header_values=["Policy Number", "Property Sec"],
        data_rows=[["POL-1", "Z"]],
    )

    assert result.contract_code_year == "BW01972"  # base only, "Z" isn't A or B
    assert result.section is None


def test_phrase_tier_takes_priority_over_property_sec_column():
    # Property Sec is a last-resort fallback, tried only after all four phrase tiers
    # come up empty - worksheet name's "Sec B" phrase wins here, and the Property Sec
    # column's "A" value is never even consulted.
    result = _detect(
        metadata_values=["Hardy"], header_values=["Policy Number", "Property Sec"],
        data_rows=[["POL-1", "A"]], worksheet_name="Sec B",
    )

    assert result.contract_code_year == "BW01972B"
    assert "worksheet_name" in result.tier


# ---------------------------------------------------------------------------
# Property Sec / Province read row by row: rows that disagree split the sheet instead
# of resolving (or dropping) a single section for the whole thing
# ---------------------------------------------------------------------------

def test_property_sec_rows_that_agree_resolve_normally_no_split():
    result = _detect(
        header_values=["Policy Number", "Property Sec"],
        data_rows=[["POL-1", "A"], ["POL-2", "A"], ["POL-3", "A"]], source_file="Hardy_Risk_test.xlsx",
    )

    assert result.contract_code_year == "BW01972A"
    assert result.row_sections is None


def test_property_sec_rows_that_disagree_produce_a_row_split():
    result = _detect(
        header_values=["Policy Number", "Property Sec"],
        data_rows=[["POL-1", "B"], ["POL-2", "A"]], source_file="Hardy_Risk_test.xlsx",
    )

    assert result.status == "resolved"
    assert result.base_code == "BW01972"
    assert result.row_sections == {6: "B", 7: "A"}  # _detect's data rows start at 6


def test_province_bc_and_alberta_map_to_a_and_b():
    result = _detect(
        header_values=["Policy Number", "Province"],
        data_rows=[["POL-1", "BC"]], source_file="Hardy_Risk_test.xlsx",
    )
    assert result.contract_code_year == "BW01972A"

    result = _detect(
        header_values=["Policy Number", "Province"],
        data_rows=[["POL-1", "Alberta"]], source_file="Hardy_Risk_test.xlsx",
    )
    assert result.contract_code_year == "BW01972B"


def test_province_only_used_when_property_sec_column_is_absent():
    # Property Sec is checked first - if present (even with a single, agreeing value),
    # Province is never consulted, even if it would have said something different
    result = _detect(
        header_values=["Policy Number", "Property Sec", "Province"],
        data_rows=[["POL-1", "A", "Alberta"]], source_file="Hardy_Risk_test.xlsx",
    )

    assert result.contract_code_year == "BW01972A"


def test_province_rows_that_disagree_produce_a_row_split():
    result = _detect(
        header_values=["Policy Number", "Province"],
        data_rows=[["POL-1", "BC"], ["POL-2", "Alberta"]], source_file="Hardy_Risk_test.xlsx",
    )

    assert result.status == "resolved"
    assert result.base_code == "BW01972"
    assert result.row_sections == {6: "A", 7: "B"}


def test_row_with_unresolvable_value_falls_back_to_none_in_the_split():
    result = _detect(
        header_values=["Policy Number", "Province"],
        data_rows=[["POL-1", "BC"], ["POL-2", "Alberta"], ["POL-3", "Ontario"]],
        source_file="Hardy_Risk_test.xlsx",
    )

    assert result.row_sections == {6: "A", 7: "B", 8: None}  # "Ontario" isn't configured


def test_no_property_sec_or_province_column_falls_through_as_before():
    result = _detect(
        header_values=["Policy Number", "Region"],
        data_rows=[["POL-1", "UK"]], source_file="Hardy_Risk_test.xlsx",
    )

    assert result.contract_code_year == "BW01972"
    assert result.row_sections is None
    assert "no section evidence found" in result.detail


# ---------------------------------------------------------------------------
# base contract unresolved: not_found / ambiguous
# ---------------------------------------------------------------------------

def test_base_not_found_when_no_tier_matches():
    # mirrors file_Claims_test.xlsx's "Section A" sheet: section-shaped sheet name,
    # but no base-contract evidence anywhere.
    result = _detect(
        metadata_values=["Client", "Example Insurance Ltd"], worksheet_name="Section A",
        source_file="file_Claims_test.xlsx",
    )

    assert result.status == "not_found"
    assert result.contract_code_year == ""
    assert "no contract evidence found" in result.detail


def test_base_ambiguous_when_a_tier_matches_two_contracts():
    result = _detect(metadata_values=["Hardy", "HDI"])

    assert result.status == "ambiguous"
    assert result.contract_code_year == ""
    assert "BW01972" in result.detail and "BW01973" in result.detail


# ---------------------------------------------------------------------------
# ambiguity reporting includes the specific cell each code was matched in
# ---------------------------------------------------------------------------

def test_ambiguous_metadata_match_reports_each_codes_cell():
    result = _detect(
        metadata_values=["Hardy", "HDI"], metadata_cells=["B1", "B3"],
    )

    assert result.status == "ambiguous"
    assert "{'BW01972': 'B1', 'BW01973': 'B3'}" in result.detail


def test_ambiguous_table_match_reports_each_codes_cell():
    # header row 5, start_col 1 (per _detect's defaults): "Hardy" in column A of the
    # header (A5), "HDI" in column B of the first data row (B6, since data starts row 6)
    result = _detect(header_values=["Claim Number", "Hardy"], data_rows=[["CLM-1", "HDI"]])

    assert result.status == "ambiguous"
    assert "{'BW01972': 'B5', 'BW01973': 'B6'}" in result.detail


def test_resolved_match_detail_includes_the_matched_cell():
    result = _detect(metadata_values=["Broker", "Hardy Ltd"], metadata_cells=["A1", "B1"])

    assert result.contract_code_year == "BW01972"
    assert "(B1)" in result.detail


def test_worksheet_name_and_filename_matches_report_a_placeholder_not_a_cell():
    ws_result = _detect(worksheet_name="CNA Claims")
    assert "(worksheet name)" in ws_result.detail

    file_result = _detect(source_file="Hardy_Risk_test.xlsx")
    assert "(filename)" in file_result.detail


# ---------------------------------------------------------------------------
# "Loss Details"/"Loss Description" columns: free-text narrative, excluded entirely
# from table evidence (header + data) so coincidental alias matches in prose don't
# produce false "multiple contracts matched" ambiguity.
# ---------------------------------------------------------------------------

def test_loss_description_column_excluded_from_evidence():
    # without the exclusion, "HDI" in the narrative column would conflict with the
    # genuine Agreement Number evidence (BW01972) and make this ambiguous
    result = _detect(
        header_values=["Agreement Number", "Loss Description"],
        data_rows=[["BW01972", "Reviewed by HDI liaison team"]],
    )

    assert result.status == "resolved"
    assert result.contract_code_year == "BW01972"


def test_loss_details_column_also_excluded():
    result = _detect(
        header_values=["Agreement Number", "Loss Details"],
        data_rows=[["BW01972", "Reviewed by HDI liaison team"]],
    )

    assert result.status == "resolved"
    assert result.contract_code_year == "BW01972"


def test_without_exclusion_this_would_have_been_ambiguous():
    # control case: same conflicting-contract narrative text, but under a column name
    # NOT on the exclusion list - proves the ambiguity is real (Agreement Number says
    # BW01972, the narrative coincidentally says HDI/BW01973) and the exclusion in the
    # tests above is what's actually preventing it, not some other difference.
    result = _detect(
        header_values=["Agreement Number", "Adjuster Notes"],
        data_rows=[["BW01972", "Reviewed by HDI liaison team"]],
    )

    assert result.status == "ambiguous"


def test_loss_description_column_name_matching_is_case_insensitive_and_trimmed():
    # column name matching is case-insensitive/trimmed, same as Property Sec elsewhere
    result = _detect(
        header_values=["Agreement Number", " loss description "],
        data_rows=[["BW01972", "Mentions HDI in passing"]],
    )

    assert result.status == "resolved"
    assert result.contract_code_year == "BW01972"


# ---------------------------------------------------------------------------
# numeric code+year pattern: digit-core + exactly two more digits.
# BW01972 -> digit-core "1972" (leading zero of "01972" dropped), so a matching
# numeric value looks like "197225" - the actual real-world case that motivated this
# rule: file1_Risk_test1.xlsx's "Broker Ref: 197225" metadata value.
# ---------------------------------------------------------------------------

def test_numeric_pattern_matches_a_pure_number_metadata_value():
    result = _detect(metadata_values=["Broker Ref", 197225])

    assert result.contract_code_year == "BW01972"
    assert result.tier == "metadata"


def test_numeric_pattern_matches_within_a_longer_text_value():
    result = _detect(metadata_values=["Ref: 197225"])

    assert result.contract_code_year == "BW01972"


def test_numeric_pattern_applies_to_filename_tier_too():
    result = _detect(source_file="Claims_197225_Aug.xlsx")

    assert result.contract_code_year == "BW01972"
    assert result.tier == "filename"


def test_numeric_pattern_requires_two_trailing_digits():
    # the digit-core alone, with nothing after it, must not match
    result = _detect(metadata_values=[1972])

    assert result.status == "not_found"


def test_numeric_pattern_does_not_match_embedded_in_a_longer_digit_run():
    # digit-core+2 sits inside a longer unbroken digit run - must not match, same
    # protection validated against the real decoy fixture ("b1262bw0197219")
    result = _detect(metadata_values=["0197225"])

    assert result.status == "not_found"
