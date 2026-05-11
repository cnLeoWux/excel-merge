## 1. Workflow service hardening

- [x] 1.1 Add a shared service-level month validation helper for `YYYYMM` with year range 2020-2099 and month range 01-12.
- [x] 1.2 Update `run_sales_report()` to reject missing or invalid `target_month` with `WorkflowError(code="usage_error", exit_code=2)` before calling `process_sales_report_workflow()`.
- [x] 1.3 Update `prepare_api_merge()` to reject invalid non-empty `month` values with `WorkflowError(code="usage_error", exit_code=2)` before calling `process_sales_report_workflow()`.
- [x] 1.4 Ensure service missing-file paths consistently raise `WorkflowError(code="file_not_found", exit_code=3)` for all operations requiring existing files.
- [x] 1.5 Ensure service write/persistence failures consistently raise `WorkflowError(code="processing_error", exit_code=4)`.
- [x] 1.6 Update API report statistics to include full workflow fields plus `report_rows`.
- [x] 1.7 Keep current empty API report behavior explicit by preserving `processing_error` on empty report data and documenting/testing it.

## 2. Adapter boundary cleanup

- [x] 2.1 Refactor `cli.py` where practical so file-not-found and workflow failure classification rely on `WorkflowError`, while preserving argparse validation, backup behavior, JSON envelope, and exit codes.
- [x] 2.2 Verify CLI invalid month still exits 2 with `usage_error` in JSON mode after service validation is added.
- [x] 2.3 Refactor `excel_merge_api.py` error mapping so `WorkflowError(code="usage_error")` maps to HTTP 400 and `processing_error` maps to HTTP 500 for both `/merge` and `/merge/json`.
- [x] 2.4 Preserve Flask-specific request validation in `excel_merge_api.py`: missing upload fields, empty filenames, extension checks for `/merge`, `secure_filename()`, and `send_file()`.
- [x] 2.5 Preserve API response shapes: attachment responses for `/merge`, API-specific `success/session_id/download_url/statistics/files` for `/merge/json`, and API-shaped failure JSON.

## 3. Documentation alignment

- [x] 3.1 Update `AGENTS.md` structure and CLI usage sections so current CLI syntax is `order_file payment_file [target_month]`, not `--month`.
- [x] 3.2 Update `AGENTS.md` parameter table to include `target_month`, `--match-only`, and `--mark-only`, and to omit `--month` as a current parameter.
- [x] 3.3 Update `AGENTS.md` Agent recommended usage to show full workflow with positional `target_month --json --quiet` and month inference/ask-before-run guidance.
- [x] 3.4 Update `documents/USAGE_EXAMPLES.md` CLI examples and argument table to use positional `target_month` and explicit reduced workflow wording.
- [x] 3.5 Review `.opencode/skills/excel-merge-cli/SKILL.md` and adjust any remaining ambiguous wording that could present `--month` as a current CLI option.

## 4. Tests

- [x] 4.1 Add unit tests for service missing-file normalization in `run_match_only()`, `run_mark_only()`, `run_sales_report()`, and/or `prepare_api_merge()` as appropriate.
- [x] 4.2 Add unit tests for service invalid month normalization in `run_sales_report()` and `prepare_api_merge()`.
- [x] 4.3 Add unit tests for service write/persistence failure normalization.
- [x] 4.4 Add unit tests for API report statistics including `marked_rows` and `report_rows`.
- [x] 4.5 Add unit tests for `prepare_api_merge()` no-month metadata.
- [x] 4.6 Add or update API integration tests for invalid `month` returning HTTP 400 and API-shaped error JSON.
- [x] 4.7 Add documentation contract tests that AGENTS.md and USAGE_EXAMPLES.md do not present `--month` as a current CLI parameter and do show positional `target_month` full-workflow examples.
- [x] 4.8 Update existing CLI/API tests only as needed to preserve current public behavior after boundary hardening.

## 5. Validation

- [x] 5.1 Run `openspec validate harden-workflow-service-boundary --strict`.
- [x] 5.2 Run `openspec validate --all --strict`.
- [x] 5.3 Run targeted pytest suites covering workflow service, CLI integration, API integration, and documentation contract tests.
- [x] 5.4 Review `workflow_service.py`, `cli.py`, and `excel_merge_api.py` to confirm workflow failures are normalized in the service and adapters mainly perform transport formatting.
