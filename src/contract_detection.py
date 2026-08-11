"""ContractCodeYear detection: identifies the base contract (and, where applicable,
its section) that an extracted table belongs to, using config/contract_codes.json.

Detection searches four evidence tiers in priority order - metadata, table column
names/values, worksheet name, source filename - stopping at the first tier that
yields a match. See CLAUDE.md-approved plan for the full rationale.

A base contract match can come from either an exact configured alias (word-boundary
isolated) or the code's digit-core followed by exactly two more digits (e.g. BW01972's
'1972' matches the '197225' in a numeric broker-ref field), isolated from any longer
digit run so it can't fire inside an unrelated bigger number.

Section matching additionally recognises a table column literally named 'Property Sec'
or 'Property Section': its first data row value is read directly as the section code,
alongside (not instead of) the phrase-based tiers.

Every match (base contract and section) is tracked back to the specific cell it came
from ('A1'-style, or '(worksheet name)'/'(filename)'/'(metadata)' where there's no
single cell) - surfaced in ContractDetectionResult.detail, including the ambiguous-
match warning, e.g. "multiple contracts matched in table_columns_and_values:
{'BW01972': 'G6', 'BW05407': 'C12'}".
"""

import json
import re
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from src.logging_utils import NULL_LOGGER


@dataclass
class ContractCodeConfig:
    contracts: dict  # canonical_code -> {"aliases": [...], "sections": {letter: [...aliases]} | None}


@dataclass
class ContractDetectionResult:
    contract_code_year: str  # e.g. "BW0197221A", "BW01973", or "" if unresolved
    status: str  # "resolved" | "not_found" | "ambiguous"
    tier: str  # evidence tier(s) that resolved it, "" if unresolved
    detail: str  # human-readable trace string, for logging + WorksheetResult.warnings
    section: str = None  # resolved section letter alone (e.g. "A"), or None if not applicable


def load_contract_codes(config_path):
    """Load and validate config/contract_codes.json into a ContractCodeConfig."""
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError(f"Contract codes config must be a JSON object: {config_path}")

    contracts = data.get("contracts")
    if not isinstance(contracts, dict):
        raise ValueError(f"Contract codes config missing 'contracts' object: {config_path}")

    for code, entry in contracts.items():
        if not isinstance(entry, dict) or not isinstance(entry.get("aliases"), list):
            raise ValueError(f"Contract '{code}' must define an 'aliases' list: {config_path}")
        sections = entry.get("sections")
        if sections is not None:
            if not isinstance(sections, dict):
                raise ValueError(f"Contract '{code}' 'sections' must be an object: {config_path}")
            for section_letter, section_entry in sections.items():
                if not isinstance(section_entry, dict) or not isinstance(section_entry.get("aliases"), list):
                    raise ValueError(
                        f"Contract '{code}' section '{section_letter}' must define an 'aliases' list: {config_path}"
                    )

    return ContractCodeConfig(contracts=contracts)


def _alias_pattern(alias):
    """Word aliases (no digits, e.g. 'Hardy') require word-boundary isolation so they
    don't match inside unrelated text. Code-style aliases (contain a digit, e.g. the
    canonical 'BW01972') are specific enough on their own that a plain substring search
    is safe - and is what lets a broker-ref style value like 'B1262BW0197219' resolve.
    """
    alias = str(alias)
    if any(c.isdigit() for c in alias):
        return re.compile(re.escape(alias), re.IGNORECASE)
    return re.compile(r"(?<![A-Za-z0-9])" + re.escape(alias) + r"(?![A-Za-z0-9])", re.IGNORECASE)


def _digit_core(code):
    """The code's digits with any leading zeros dropped, e.g. 'BW01972' -> '1972'."""
    return re.sub(r"\D", "", str(code)).lstrip("0")


def _numeric_year_pattern(code):
    """Match the code's digit-core immediately followed by exactly two more digits
    (e.g. 1972 -> matches the '197225' in a 'Broker Ref: 197225' cell), isolated from
    any longer digit run so it can't match inside an unrelated bigger number. None if
    the code has too few digits to be a specific enough fragment (avoids over-broad
    matches like any 2-digit number in the workbook).
    """
    digit_core = _digit_core(code)
    if len(digit_core) < 3:
        return None
    return re.compile(r"(?<!\d)" + re.escape(digit_core) + r"\d{2}(?!\d)")


def _build_alias_index(aliases_by_code):
    """aliases_by_code: {code: [alias, ...]} -> [(compiled_pattern, code), ...]"""
    return [(_alias_pattern(alias), code) for code, aliases in aliases_by_code.items() for alias in aliases]


def _build_contract_index(contracts):
    """Base-contract matching: exact alias match OR the code's digit-core+year pattern,
    e.g. BW01972's aliases plus the '1972NN' numeric pattern.
    """
    index = []
    for code, entry in contracts.items():
        for alias in entry["aliases"]:
            index.append((_alias_pattern(alias), code))
        numeric_pattern = _numeric_year_pattern(code)
        if numeric_pattern is not None:
            index.append((numeric_pattern, code))
    return index


def _cell_ref(col, row):
    """'A1'-style reference, or a generic placeholder if position info isn't known."""
    if col is None or row is None:
        return "(unknown cell)"
    return f"{get_column_letter(col)}{row}"


def _find_alias_matches(alias_index, evidence):
    """evidence: [(text, cell_ref), ...]. Returns {code: cell_ref} - the first cell
    where each code was matched, for traceability (e.g. in the ambiguous-match warning).
    """
    matched = {}
    for text, cell_ref in evidence:
        for pattern, code in alias_index:
            if code not in matched and pattern.search(text):
                matched[code] = cell_ref
    return matched


_SECTION_COLUMN_NAMES = {"property sec", "property section"}


def _section_from_property_column(header_values, data_rows, sections, start_col, data_row_numbers):
    """If a header column is literally named 'Property Sec'/'Property Section'
    (case-insensitive), read the section letter directly from that column's first data
    row value - the value itself IS the section code, not a phrase to search for. None
    if there's no such column, no data, or the value doesn't match a configured section.
    Returns (section, cell_ref) on a match.
    """
    for idx, header in enumerate(header_values):
        if header is not None and str(header).strip().lower() in _SECTION_COLUMN_NAMES:
            if not data_rows or idx >= len(data_rows[0]) or data_rows[0][idx] is None:
                return None
            value = str(data_rows[0][idx]).strip()
            if value not in sections:
                return None
            row = data_row_numbers[0] if data_row_numbers else None
            col = (start_col + idx) if start_col is not None else None
            return value, _cell_ref(col, row)
    return None


def _metadata_evidence(metadata):
    cells = metadata.cells if len(metadata.cells) == len(metadata.values) else [None] * len(metadata.values)
    return [(str(v), cell or "(metadata)") for v, cell in zip(metadata.values, cells)]


def _table_evidence(header_values, data_rows, header_row, start_col, data_row_numbers):
    evidence = []
    for idx, v in enumerate(header_values):
        if v is not None:
            col = (start_col + idx) if start_col is not None else None
            evidence.append((str(v), _cell_ref(col, header_row)))

    row_numbers = data_row_numbers if len(data_row_numbers) == len(data_rows) else [None] * len(data_rows)
    for row, row_num in zip(data_rows, row_numbers):
        for idx, v in enumerate(row):
            if v is not None:
                col = (start_col + idx) if start_col is not None else None
                evidence.append((str(v), _cell_ref(col, row_num)))
    return evidence


def _evidence_tiers(source_file, worksheet_name, metadata, header_values, data_rows, header_row, start_col, data_row_numbers):
    return [
        ("metadata", _metadata_evidence(metadata)),
        ("table_columns_and_values", _table_evidence(header_values, data_rows, header_row, start_col, data_row_numbers)),
        ("worksheet_name", [(worksheet_name, "(worksheet name)")]),
        ("filename", [(source_file.stem, "(filename)")]),
    ]


def detect_contract_code_year(
    *, source_file, worksheet_name, metadata, header_values, data_rows,
    header_row, start_col, data_row_numbers,
    contract_config, logger=NULL_LOGGER, ingestion_type=None,
):
    tiers = _evidence_tiers(
        source_file, worksheet_name, metadata, header_values, data_rows, header_row, start_col, data_row_numbers,
    )

    contract_alias_index = _build_contract_index(contract_config.contracts)

    code = None
    base_tier = ""
    base_cell = ""
    for tier_name, evidence in tiers:
        matched = _find_alias_matches(contract_alias_index, evidence)
        if len(matched) == 1:
            code = next(iter(matched))
            base_tier = tier_name
            base_cell = matched[code]
            break
        if len(matched) > 1:
            return ContractDetectionResult(
                contract_code_year="", status="ambiguous", tier=tier_name,
                detail=f"multiple contracts matched in {tier_name}: {matched}",
            )

    if code is None:
        return ContractDetectionResult(
            contract_code_year="", status="not_found", tier="",
            detail="no contract evidence found in metadata, table, worksheet name, or filename",
        )

    sections = contract_config.contracts[code].get("sections")
    if not sections:
        return ContractDetectionResult(
            contract_code_year=code, status="resolved", tier=base_tier,
            detail=f"matched {code} via {base_tier} ({base_cell})",
        )

    section_alias_index = _build_alias_index(
        {letter: entry["aliases"] for letter, entry in sections.items()}
    )
    property_column_match = _section_from_property_column(header_values, data_rows, sections, start_col, data_row_numbers)

    for tier_name, evidence in tiers:
        matched_sections = _find_alias_matches(section_alias_index, evidence)
        if tier_name == "table_columns_and_values" and property_column_match is not None:
            prop_section, prop_cell = property_column_match
            matched_sections.setdefault(prop_section, prop_cell)
        if len(matched_sections) == 1:
            section = next(iter(matched_sections))
            return ContractDetectionResult(
                contract_code_year=f"{code}{section}", status="resolved", tier=f"{base_tier}+{tier_name}",
                detail=(
                    f"matched {code} via {base_tier} ({base_cell}); "
                    f"section {section} via {tier_name} ({matched_sections[section]})"
                ),
                section=section,
            )
        if len(matched_sections) > 1:
            return ContractDetectionResult(
                contract_code_year=code, status="resolved", tier=base_tier,
                detail=(
                    f"matched {code} via {base_tier} ({base_cell}); section ambiguous in {tier_name}: "
                    f"{matched_sections} - section dropped"
                ),
            )

    return ContractDetectionResult(
        contract_code_year=code, status="resolved", tier=base_tier,
        detail=f"matched {code} via {base_tier} ({base_cell}); no section evidence found",
    )
