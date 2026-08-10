
## 1. Project Purpose

Build a controlled, repeatable Python workflow for ingesting Claims and Risk Excel files.

The workflow should locate tabular data within inconsistent source workbooks, extract it into a clean standard structure, preserve relevant metadata, validate outputs, and maintain processing evidence.

---

## 2. Core Processing Requirements

The workflow must:

1. Read source file paths from `input/filelist.csv`. Use the `Type` column to determine whether each file is processed as `Risk` or `Claims`.
2. Check that each source file exists and can be accessed.
3. Use `config/password.json` when a password-protected file requires access.
4. Inspect each relevant worksheet in the source workbook.
5. Detect table headers using the appropriate Risk or Claims anchor pairs defined in `config/anchor_pairs.json`.
6. Extract the detected header and associated data below it.
7. Extract relevant content above the detected header as metadata.
8. Skip worksheets where no valid header is detected.
9. Skip detected tables containing no meaningful data rows.
10. Write each clean processed table starting at cell A1.
11. Keep Risk and Claims outputs separated.
12. Validate processing results against agreed validation rules.
13. Record processing outcomes, validation results, warnings and failures.
14. Never modify the original source files.

---

## 3. Architecture

Use a modular project structure:

    main.py
    requirements.txt

    src/
        file_access.py
        anchor_detection.py
        metadata_extraction.py
        sheet_processing.py
        validation.py
        logging_utils.py

    config/
        anchor_pairs.json
        password.json

    input/
    output/
        risk/
        claims/

    tests/
    logs/

`main.py` is the workflow orchestrator.

Keep reusable processing logic in `/src`.

Keep configurable business rules outside Python where practical.

Avoid unnecessary architectural complexity.

---

## 4. Development Principles

Prioritise:

- simplicity
- modularity
- readability
- deterministic behaviour
- testability
- traceability
- reproducibility

Functions should have clear responsibilities and predictable inputs and outputs.

Do not duplicate logic or introduce special-case fixes when a general solution is appropriate.

---

## 5. Testing

Use `pytest`.

Test individual processing components before relying on them in the end-to-end workflow.

Use the workbooks under `/tests/test files` as controlled test cases.

When a test fails:

1. reproduce the failure;
2. identify why the current logic failed;
3. determine whether it represents a legitimate missing requirement;
4. propose a generalisable change;
5. implement the smallest appropriate change;
6. rerun all relevant tests to check for regression.

Do not create filename-specific fixes merely to make a test pass.

---

## 6. Validation and Logging

Successful code execution does not automatically mean successful ingestion.

Outputs must be validated against agreed validation rules.

Record sufficient evidence to understand each processing run, including where appropriate:

- source file;
- ingestion type;
- worksheet;
- anchor detection result;
- detected header;
- rows and columns extracted;
- validation result;
- output location;
- warnings;
- errors.

Never log passwords or secrets.

---

## 7. Collaboration Principles

Act as a development collaborator rather than independently redesigning the project.

Before material changes:

1. inspect the relevant existing files;
2. understand the requirement;
3. identify the responsible module;
4. explain important assumptions;
5. propose the approach;
6. identify required tests;
7. wait for approval when the change materially affects architecture or processing behaviour.

Prefer incremental changes.

Do not rewrite unrelated working code.

Preserve existing passing behaviour unless a requirement explicitly changes it.

---

## 8. Definition of Done

A material feature is complete when:

- the requirement is understood;
- implementation is in the appropriate module;
- relevant tests exist and pass;
- validation behaviour is defined where required;
- meaningful processing evidence is recorded;
- existing functionality has not unintentionally regressed.

The objective is not simply to make the script run.

The workflow should remain understandable, testable, validated, reproducible and safely changeable.