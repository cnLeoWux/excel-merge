## Why

CLI、Agent Skill、交互入口与 HTTP API 目前暴露了多套相近但不一致的契约：CLI 使用 `ok/data/error` JSON 信封，API 使用 `success/download_url/files` 形状；CLI 默认完整流程依赖 `target_month`，而文档与自动化调用方式容易误解为可默认基础匹配。现在先对齐这些入口契约，避免后续抽 workflow/service 层或拆分 `utils.py` 时把不一致固化到更多模块。

## What Changes

- Standardize the default automation intent: two uploaded/provided files should run the full sales-report workflow when `target_month` is available or can be obtained; matching-only is an explicit reduced workflow.
- Clarify how Agents and Skills acquire `target_month`: infer from filenames/conversation when reliable, otherwise ask the user before invoking the CLI.
- Align CLI JSON and documentation around the current `target_month --json --quiet` invocation for the full workflow, including `marked_rows` in full-workflow statistics.
- Decide and document the relationship between CLI JSON and HTTP API JSON rather than leaving `/merge/json` and CLI as contradictory contracts.
- Align HTTP API request/response behavior with the chosen contract, including file validation, error shape, sales-report behavior, and downloadable artifacts.
- Keep CLI and API file-output semantics explicit: CLI writes the order file in place and does not produce `report_*.xlsx`; API may persist downloadable result/report files under `results/` if that remains the selected API contract.
- Update Agent-facing documentation and tests so they validate the same CLI/API contracts instead of preserving older contradictory assumptions.

## Capabilities

### New Capabilities

- None.

### Modified Capabilities

- `cli-input`: define the standard full-workflow invocation, `target_month` acquisition expectations, and explicit reduced-mode behavior.
- `cli-output`: align JSON statistics, stdout/stderr expectations, cancellation behavior, and error codes with the selected CLI contract.
- `http-api`: align `/merge`, `/merge/json`, error responses, file validation, MIME/download behavior, and sales-report artifacts with the selected API contract.
- `agent-documentation`: update Agent/Skill-facing usage rules so automation obtains a month for the default full workflow and does not silently fall back to matching-only.
- `automated-testing`: update expected CLI/API test behavior to validate the aligned contracts.
- `sales-report`: clarify how CLI, interactive mode, and API trigger the full workflow and how API-specific report persistence relates to the core workflow.

## Impact

- Affected entry points: `cli.py`, `excel_merge.py`, `excel_merge_api.py`.
- Affected Agent surfaces: `AGENTS.md`, `.opencode/skills/excel-merge-cli/SKILL.md`, and user-facing usage documents.
- Affected tests: CLI subprocess/main tests, Flask API integration tests, and any tests asserting JSON shape or report-file behavior.
- Affected specs: `cli-input`, `cli-output`, `http-api`, `agent-documentation`, `automated-testing`, and `sales-report`.
- No new runtime dependency is expected; changes should focus on contract alignment and compatibility decisions before broader workflow/service refactoring.
