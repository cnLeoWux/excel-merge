## MODIFIED Requirements

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
- **THEN** there MUST be tests for `cli.py` that execute the script as a subprocess or call `main_cli()` with patched `sys.argv` and verify:
  - Default full workflow with positional `target_month` writes sales-report markings back in place to the order file and produces no `report_*.xlsx` artefact.
  - Agent/Skill documentation requires month inference or user prompt before default full-workflow invocation.
  - Matching-only behavior is covered as an explicit reduced workflow using `--match-only`.
  - Passing the removed flags `-o`/`--output`/`--output-dir` yields exit code 2.
  - The `--json` output is a valid JSON with the expected structure for the executed mode.
  - Invalid arguments cause a non-zero exit code.

#### Scenario: API integration tests
- **WHEN** running the test suite
- **THEN** there MUST be integration tests for `excel_merge_api.py` that:
  - Start the Flask test server.
  - Send `POST` requests to `/merge` and `/merge/json` with and without the `month` parameter.
  - Verify `/merge/json` returns the documented API-specific `success` response shape.
  - Verify the file content or JSON response is correct.
  - Verify the `/health` endpoint returns a 200 status code.
