## MODIFIED Requirements

### Requirement: Test Coverage

The test suite MUST provide comprehensive coverage for core logic, workflow/service orchestration, entry points, documentation contracts, and API endpoints.

#### Scenario: Core logic unit tests
- **WHEN** running the test suite
- **THEN** there MUST be unit tests for `utils.py` that cover:
  - `process_excel_files` matching logic (exact, P-number, hyphen).
  - `add_sales_report_period` marking logic ("全退", "已取消").
  - `filter_unmarked_and_generate_report` filtering logic.
  - `read_file_with_appropriate_method` for all supported file types and encodings.

#### Scenario: Workflow service unit tests
- **WHEN** running the test suite
- **THEN** there MUST be unit tests for the workflow/service layer that cover:
  - Match-only service operation and statistics.
  - Mark-only service operation and statistics.
  - Full sales-report service operation and statistics.
  - API-oriented service result metadata for downloadable artifacts in both month and no-month modes.
  - Service-level error normalization for missing files, invalid months, and write/persistence failures.
  - API report statistics including `report_rows` and full workflow statistics.

#### Scenario: CLI functional tests
- **WHEN** running the test suite
- **THEN** there MUST be tests for `cli.py` that execute the script as a subprocess or call `main_cli()` with patched `sys.argv` and verify:
  - Basic matching or match-only mode writes the result back in place to the order file (no separate output file is produced).
  - The sales report workflow (positional `target_month`) writes 销售报表 markings back in place to the order file and produces no `report_*.xlsx` artefact.
  - CLI output remains compatible after execution is routed through the workflow/service layer.
  - Passing the removed flags `-o`/`--output`/`--output-dir` yields exit code 2.
  - The `--json` output is a valid JSON with the expected structure for the executed mode.
  - Invalid arguments and service usage errors cause a non-zero exit code with the documented error envelope.

#### Scenario: API integration tests
- **WHEN** running the test suite
- **THEN** there MUST be integration tests for `excel_merge_api.py` that:
  - Start the Flask test server.
  - Send `POST` requests to `/merge` and `/merge/json` with and without the `month` parameter.
  - Verify API responses remain compatible after execution is routed through the workflow/service layer.
  - Verify invalid API `month` values return HTTP 400 with API-shaped errors.
  - Verify the file content or JSON response is correct.
  - Verify the `/health` endpoint returns a 200 status code.

#### Scenario: Documentation contract tests
- **WHEN** running the test suite
- **THEN** there MUST be tests or assertions that Agent/user-facing docs do not present `--month` as a current CLI parameter
- **AND** docs MUST show positional `target_month` for full workflow examples
- **AND** docs MUST describe `--match-only` as an explicit reduced workflow only
