## 1. Workflow service module

- [x] 1.1 Create `workflow_service.py` with lightweight `WorkflowResult`, `WorkflowError`, and `ApiWorkflowResult` dataclasses.
- [x] 1.2 Add shared statistics helpers for match statistics, mark statistics, full workflow statistics, and API report statistics.
- [x] 1.3 Implement `run_match_only(order_file, payment_file, *, verbose=False, write_back=True)` using existing `utils.process_excel_files()` and `write_result_file()`.
- [x] 1.4 Implement `run_mark_only(order_file, *, verbose=False, write_back=True)` using existing file reading, `add_sales_report_period()`, and `write_result_file()`.
- [x] 1.5 Implement `run_sales_report(order_file, payment_file, target_month, *, verbose=False, write_back=True)` using existing `process_sales_report_workflow()` and in-place write-back semantics.
- [x] 1.6 Implement API-oriented workflow helper(s) that prepare result paths, download names, download URLs, statistics, and file metadata while preserving API downloadable artifact behavior.
- [x] 1.7 Add normalized service error handling for known file-not-found, usage, and processing failures without changing adapter-specific output formats.

## 2. CLI adapter refactor

- [x] 2.1 Refactor `cli.py` match-only branch to call `run_match_only()` and format the returned result.
- [x] 2.2 Refactor `cli.py` mark-only branch to call `run_mark_only()` and format the returned result.
- [x] 2.3 Refactor `cli.py` full-workflow branch to call `run_sales_report()` and format the returned result.
- [x] 2.4 Preserve CLI argument parsing, target-month validation, interactive prompting, logging setup, JSON envelope, text output, and exit codes.
- [x] 2.5 Confirm CLI still creates backups according to the current CLI contract or document any service-layer ownership decision before implementation.

## 3. Interactive adapter refactor

- [x] 3.1 Refactor `excel_merge.py` full sales-report path to call `run_sales_report()`.
- [x] 3.2 Refactor `excel_merge.py` basic processing path to call `run_match_only()` or the appropriate service operation.
- [x] 3.3 Preserve interactive file selection, non-interactive argument handling, JSON output support, and exit code behavior.

## 4. HTTP API adapter refactor

- [x] 4.1 Refactor `/merge` without `month` to call the API-oriented service helper and return the generated attachment.
- [x] 4.2 Refactor `/merge` with `month` to call the API-oriented service helper and return the report attachment when report data is produced.
- [x] 4.3 Refactor `/merge/json` without `month` to call the API-oriented service helper and return the documented API-specific `success` response shape.
- [x] 4.4 Refactor `/merge/json` with `month` to call the API-oriented service helper and return `statistics.report_rows` plus download metadata.
- [x] 4.5 Preserve Flask-specific concerns in `excel_merge_api.py`: upload parsing, required-field validation, `secure_filename()`, HTTP status codes, and `send_file()`.

## 5. Tests

- [x] 5.1 Add unit tests for service statistics helpers.
- [x] 5.2 Add unit tests for `run_match_only()`, `run_mark_only()`, and `run_sales_report()` using temporary files/fixtures.
- [x] 5.3 Add unit tests for API-oriented service result metadata, including result path, download URL, files metadata, and `report_rows`.
- [x] 5.4 Update CLI tests to verify behavior remains compatible after routing execution through the service layer.
- [x] 5.5 Update API integration tests to verify `/merge` and `/merge/json` remain compatible after routing execution through the service layer.
- [x] 5.6 Ensure tests verify no CLI `report_*.xlsx` artifact is created for full workflow.

## 6. Validation

- [x] 6.1 Run `openspec validate introduce-workflow-service-layer --strict`.
- [x] 6.2 Run `openspec validate --all --strict`.
- [x] 6.3 Run the relevant pytest suite, including workflow service tests, CLI tests, and API integration tests.
- [x] 6.4 Review `cli.py`, `excel_merge.py`, and `excel_merge_api.py` to confirm they are thinner adapters and no longer duplicate shared statistics formulas.
