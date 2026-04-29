## Purpose

自动化测试能力 - 定义项目使用 `pytest` 作为测试框架的结构、覆盖范围与执行约定。该能力由 `tests/` 目录、`requirements-dev.txt` 与 `tests/conftest.py` 共享夹具实现，覆盖 `utils.py` 核心逻辑、`cli.py` 子进程入口与 `excel_merge_api.py` Flask 端点。

## Requirements

### Requirement: Test Framework and Structure
The project MUST use `pytest` as its testing framework. All test files MUST be located in a new top-level `tests/` directory.

#### Scenario: Test file structure
- **WHEN** new tests are added
- **THEN** they MUST be placed in files named `test_*.py` or `*_test.py` inside the `tests/` directory.

#### Scenario: Development dependencies
- **WHEN** setting up the development environment
- **THEN** a `requirements-dev.txt` file MUST exist and contain `pytest`.

### Requirement: Test Coverage
The test suite MUST provide comprehensive coverage for core logic, entry points, and API endpoints.

#### Scenario: Core logic unit tests
- **WHEN** running the test suite
- **THEN** there MUST be unit tests for `utils.py` that cover:
  - `process_excel_files` matching logic (exact, P-number, hyphen).
  - `add_sales_report_period` marking logic ("全退", "已取消").
  - `filter_unmarked_and_generate_report` filtering logic.
  - `read_file_with_appropriate_method` for all supported file types and encodings.

#### Scenario: CLI functional tests
- **WHEN** running the test suite
- **THEN** there MUST be tests for `cli.py` that execute the script as a subprocess and verify:
  - Basic matching writes the result back in place to the order file (no separate output file is produced).
  - The sales report workflow (`--month`) writes 销售报表 markings back in place to the order file and produces no `report_*.xlsx` artefact.
  - Passing the removed flags `-o`/`--output`/`--output-dir` yields exit code 2.
  - The `--json` output is a valid JSON with the expected structure (`data` containing only `output_file` and `statistics`).
  - Invalid arguments cause a non-zero exit code.

#### Scenario: API integration tests
- **WHEN** running the test suite
- **THEN** there MUST be integration tests for `excel_merge_api.py` that:
  - Start the Flask test server.
  - Send `POST` requests to `/merge` and `/merge/json` with and without the `month` parameter.
  - Verify the file content or JSON response is correct.
  - Verify the `/health` endpoint returns a 200 status code.

### Requirement: Test Execution and Documentation
The test suite MUST be easy to discover and run.

#### Scenario: Running tests
- **WHEN** a developer runs `pytest` from the project root
- **THEN** all automated tests MUST be discovered and executed.

#### Scenario: README documentation
- **WHEN** a developer reads `README.md`
- **THEN** it MUST contain a section explaining how to install development dependencies and run the test suite.
