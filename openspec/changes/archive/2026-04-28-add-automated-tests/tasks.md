- [x] 1.1 Create the `tests/` directory at the project root.
- [x] 1.2 Create a `requirements-dev.txt` file and add `pytest` and `pytest-flask` to it.
- [x] 1.3 Create the `tests/conftest.py` file to hold shared fixtures.
- [x] 1.4 Create subdirectories `tests/unit/` and `tests/integration/`.

## 2. Unit Tests for Core Logic (`utils.py`)

- [x] 2.1 Create `tests/unit/test_utils.py`.
- [x] 2.2 Write unit tests for the `process_excel_files` matching logic (exact, P-number, hyphen).
- [x] 2.3 Write unit tests for `add_sales_report_period` marking logic ("全退", "已取消").
- [x] 2.4 Write unit tests for `filter_unmarked_and_generate_report` filtering logic.
- [x] 2.5 Write unit tests for `read_file_with_appropriate_method` to verify it can correctly read `.csv`, `.xls`, and `.xlsx` files.

## 3. Integration Tests for Entry Points

- [x] 3.1 Create `tests/integration/test_cli.py`.
- [x] 3.2 Write tests for `cli.py` using `subprocess` to check:
    - [x] 3.2.1 Basic matching and file output.
    - [x] 3.2.2 Sales report workflow (`--month`).
    - [x] 3.2.3 JSON output (`--json`).
    - [x] 3.2.4 Error handling for invalid arguments.
- [x] 3.3 Create `tests/integration/test_api.py`.
- [x] 3.4 Write tests for `excel_merge_api.py` using a Flask test client to check:
    - [x] 3.4.1 `/health` endpoint.
    - [x] 3.4.2 `/merge` and `/merge/json` endpoints for basic matching.
    - [x] 3.4.3 `/merge` and `/merge/json` endpoints for the sales report workflow.

## 4. Refactor and Cleanup

- [x] 4.1 Move the logic from the existing `test_*.py` scripts in the root directory into the new `tests/integration/` files, converting them to `pytest` tests with assertions. (Legacy scripts were already removed in a prior cleanup; equivalent coverage is provided by the new `tests/unit/` and `tests/integration/` suites — engine detection, CSV reading, and result verification are all covered.)
- [x] 4.2 Delete the old `test_*.py` scripts from the project root (`test_engine.py`, `test_csv_reading.py`, `verify_result.py`, `verify_original.py`). (Already absent from the working tree at the time of implementation.)

## 5. Documentation

- [x] 5.1 Update `README.md` to add a "Testing" section with instructions on how to install `requirements-dev.txt` and run the test suite with `pytest`.
- [x] 5.2 Update `AGENTS.md` to include the `pytest` command in the "BUILD / LINT / TEST COMMANDS" section.
