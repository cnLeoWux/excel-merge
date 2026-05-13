## Purpose

自动化测试能力 - 定义项目使用 `pytest` 作为测试框架的结构、覆盖范围与执行约定。该能力由 `tests/` 目录、`requirements-dev.txt` 与 `tests/conftest.py` 共享夹具实现，覆盖 core modules、`utils.py` compatibility facade、`cli.py` 子进程入口与 `excel_merge_api.py` Flask 端点。

## Requirements

### Requirement: 测试框架与结构
项目 MUST 使用 `pytest` 作为测试框架。所有测试文件 MUST 位于新建的顶层 `tests/` 目录中。

#### Scenario: 测试文件结构
- **WHEN** 新测试被添加
- **THEN** 它们 MUST 放置在 `tests/` 目录内名为 `test_*.py` 或 `*_test.py` 的文件中。

#### Scenario: 开发依赖
- **WHEN** 搭建开发环境
- **THEN** 必须存在 `requirements-dev.txt` 文件且包含 `pytest`。

### Requirement: 测试覆盖
测试套件 MUST 对核心逻辑、workflow/service 编排、入口点和 API 端点提供全面覆盖。

重构 core modules 之前 SHALL 先有 behaviour-locking tests，用于锁定 matching priority、file I/O fallback 和 adapter output contracts。

#### Scenario: 核心逻辑单元测试
- **WHEN** 运行测试套件
- **THEN** MUST 存在覆盖 `utils.py` 的单元测试，包含：
  - `process_excel_files` matching logic (exact, P-number, hyphen).
  - fallback payment-file row order where an earlier hyphen match can beat a later P-number match.
  - positive orders, refund orders, and zero-amount orders.
  - `process_excel_files` preserving the current `销售报表账期` refresh side effect.
  - `add_sales_report_period` marking logic ("全退", "已取消").
  - `filter_unmarked_and_generate_report` filtering logic.
  - `read_file_with_appropriate_method` for all supported file types and encodings.

#### Scenario: Workflow service 单元测试
- **WHEN** 运行测试套件
- **THEN** MUST 存在覆盖 workflow/service 层的单元测试，包含：
  - Match-only service operation and statistics.
  - Mark-only service operation and statistics.
  - Full sales-report service operation and statistics.
  - API-oriented service result metadata for downloadable artifacts in both month and no-month modes.
  - Service-level error normalization for missing files, invalid months, and write/persistence failures.
  - API report statistics including `report_rows` and full workflow statistics.
  - Service statistics being consumed by adapters without duplicating shared formulas.

#### Scenario: CLI 功能测试
- **WHEN** 运行测试套件
- **THEN** MUST 存在针对 `cli.py` 的测试，可通过子进程执行脚本或使用补丁后的 `sys.argv` 调用 `main_cli()` 并验证：
  - Basic matching or match-only mode writes the result back in place to the order file (no separate output file is produced).
  - The sales report workflow (positional `target_month`) writes 销售报表 markings back in place to the order file and produces no `report_*.xlsx` artefact.
  - CLI output remains compatible after execution is routed through the workflow/service layer.
  - Passing the removed flags `-o`/`--output`/`--output-dir` yields exit code 2.
  - The `--json` output is a valid JSON with the expected structure for the executed mode.
  - Invalid arguments cause a non-zero exit code.

#### Scenario: API 集成测试
- **WHEN** 运行测试套件
- **THEN** MUST 存在针对 `excel_merge_api.py` 的集成测试，且：
  - Start the Flask test server.
  - Send `POST` requests to `/merge` and `/merge/json` with and without the `month` parameter.
  - Verify API responses remain compatible after execution is routed through the workflow/service layer.
  - Verify invalid API `month` values return HTTP 400 with API-shaped errors.
  - Verify the file content or JSON response is correct.
  - Verify the `/health` endpoint returns a 200 status code.

#### Scenario: 文档契约测试
- **WHEN** 运行测试套件
- **THEN** MUST 存在测试或断言，确保面向 Agent/用户的文档不会将 `--month` 表示为当前 CLI 参数
- **AND** 文档 MUST 在完整工作流示例中展示位置参数 `target_month`
- **AND** 文档 MUST 将 `--match-only` 描述为仅用于显式降级工作流

### Requirement: Refactor safety tests

行为保持型 refactors MUST 包含 tests，以便在 module extraction 或 helper rewrites 合并前，让意外的 contract changes 可见。

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

### Requirement: 测试执行与文档
测试套件 MUST 易于发现和运行。

#### Scenario: 运行测试
- **WHEN** a developer runs `pytest` from the project root
- **THEN** all automated tests MUST be discovered and executed.

#### Scenario: README 文档
- **WHEN** a developer reads `README.md`
- **THEN** it MUST contain a section explaining how to install development dependencies and run the test suite.
