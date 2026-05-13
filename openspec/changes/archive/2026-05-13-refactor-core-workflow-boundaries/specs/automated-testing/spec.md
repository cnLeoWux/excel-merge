## MODIFIED Requirements

### Requirement: Test Coverage

test suite MUST 为 core logic、workflow/service orchestration、entry points 和 API endpoints 提供全面覆盖。重构 core modules 之前 SHALL 先有 behaviour-locking tests，用于锁定 matching priority、file I/O fallback 和 adapter output contracts。

#### Scenario: Core logic unit tests
- **WHEN** 运行 test suite 时
- **THEN** 必须有覆盖以下内容的 core logic unit tests：
  - `process_excel_files` matching logic (exact, P-number, hyphen).
  - fallback payment-file row order where an earlier hyphen match can beat a later P-number match.
  - positive orders, refund orders, and zero-amount orders.
  - `process_excel_files` preserving the current `销售报表账期` refresh side effect.
  - `add_sales_report_period` marking logic (`"全退"`, `"已取消"`).
  - `filter_unmarked_and_generate_report` filtering logic.
  - `read_file_with_appropriate_method` for all supported file types and encodings.

#### Scenario: Workflow service unit tests
- **WHEN** 运行 test suite 时
- **THEN** 必须有覆盖 workflow/service layer 的 unit tests：
  - Match-only service operation and statistics.
  - Mark-only service operation and statistics.
  - Full sales-report service operation and statistics.
  - API-oriented service result metadata for downloadable artifacts in both month and no-month modes.
  - Service-level error normalization for missing files, invalid months, and write/persistence failures.
  - API report statistics including `report_rows` and full workflow statistics.
  - Service statistics being consumed by adapters without duplicating shared formulas.

#### Scenario: CLI functional tests
- **WHEN** 运行 test suite 时
- **THEN** 必须有针对 `cli.py` 的 tests，它们以 subprocess 方式执行脚本或在 patch 后的 `sys.argv` 下调用 `main_cli()` 并验证：
  - Basic matching or match-only mode writes the result back in place to the order file (no separate output file is produced).
  - The sales report workflow (positional `target_month`) writes 销售报表 markings back in place to the order file and produces no `report_*.xlsx` artefact.
  - CLI output remains compatible after execution is routed through the workflow/service layer.
  - Passing the removed flags `-o`/`--output`/`--output-dir` yields exit code 2.
  - The `--json` output is a valid JSON with the expected structure for the executed mode.
  - Invalid arguments cause a non-zero exit code.

#### Scenario: API integration tests
- **WHEN** 运行 test suite 时
- **THEN** 必须有针对 `excel_merge_api.py` 的 integration tests，它们：
  - Start the Flask test server.
  - Send `POST` requests to `/merge` and `/merge/json` with and without the `month` parameter.
  - Verify API responses remain compatible after execution is routed through the workflow/service layer.
  - Verify invalid API `month` values return HTTP 400 with API-shaped errors.
  - Verify the file content or JSON response is correct.
  - Verify the `/health` endpoint returns a 200 status code.

## ADDED Requirements

### Requirement: Refactor safety tests

行为保持型 refactors MUST 包含 tests，以便在 module extraction 或 helper rewrites 合并前，让意外的 contract changes 可见。

说明：这些 tests 是重构安全网，不要求测试内部 helper 名称。优先断言输入/输出、写回文件、错误码、stdout/stderr/API response 等外部可观察行为；helper 级测试只用于覆盖难以通过端到端 fixture 精准定位的业务分支。

#### Scenario: Matching priority golden test
- **WHEN** 某个 test dataset 同时包含 no exact match、较早的 valid hyphen fallback candidate 和较晚的 valid P-number fallback candidate 时
- **THEN** 在当前 contract 下，expected matched payment row SHALL 是较早的 hyphen candidate
- **AND** 如果实现被改成 global P-number-before-hyphen priority，该 test SHALL 失败

#### Scenario: File I/O fallback golden test
- **WHEN** test fixtures 覆盖 CSV encoding fallback、separator fallback、comment-line skipping 和 identifier cleanup 时
- **THEN** 在将 I/O code 移入专用 module 前后，expected DataFrame values SHALL 保持一致

#### Scenario: Compatibility facade test
- **WHEN** tests 从 `utils.py` 导入 core functions 时
- **THEN** matching、file I/O 和 sales-report functions 的 imports SHALL 继续成功
- **AND** 这些 functions SHALL 委派给重构后的实现且不改变 observable behaviour

#### Scenario: Adapter contract regression test
- **WHEN** core module extraction 之后运行 CLI 和 API tests 时
- **THEN** CLI JSON SHALL 保留 `ok/data/error` envelope
- **AND** API JSON SHALL 保留现有 API-specific shape
- **AND** 两个 adapter 都 SHALL NOT 在 user-facing responses 中暴露 internal module names
