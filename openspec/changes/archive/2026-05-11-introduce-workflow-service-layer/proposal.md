## Why

CLI、交互入口和 HTTP API 目前各自编排文件读取、核心处理、统计、写回、错误输出和下载响应，导致相同行为在多个入口中重复且容易漂移。引入统一 workflow/service 层可以先稳定入口与业务编排边界，为后续拆分 `utils.py`、纯化匹配引擎和改善持久化策略降低风险。

## What Changes

- Add a workflow/service layer that exposes stable application-level operations for matching and sales-report workflows.
- Move shared orchestration concerns out of entry points: statistics building, normalized result objects, normalized error objects, and write-back coordination.
- Refactor `cli.py`, `excel_merge.py`, and `excel_merge_api.py` to call the workflow/service layer instead of duplicating orchestration logic.
- Preserve current behavior and public contracts from the aligned CLI/API specs, including full-workflow default semantics, CLI in-place writes, and API downloadable result/report files.
- Keep `utils.py` business functions available and compatible; this change wraps and coordinates them rather than splitting or rewriting matching logic.
- Add/adjust tests so each entry point verifies adapter behavior while shared workflow behavior is tested through the service layer.

## Capabilities

### New Capabilities

- `workflow-service`: Defines the application service layer that coordinates core file processing, sales-report workflow execution, result statistics, persistence decisions, and normalized success/error results for entry points.

### Modified Capabilities

- `cli-input`: CLI behavior remains contract-compatible but routes execution through the workflow/service layer.
- `cli-output`: CLI JSON/text output remains contract-compatible while using workflow/service result objects for statistics and error mapping.
- `http-api`: API behavior remains contract-compatible while using workflow/service operations for matching and sales-report processing.
- `sales-report`: Sales-report workflow remains semantically unchanged but is invoked through the workflow/service layer by entry points.
- `automated-testing`: Tests must cover the new workflow/service layer and ensure entry points remain thin adapters.

## Impact

- Affected code: new workflow/service module, `cli.py`, `excel_merge.py`, `excel_merge_api.py`, and possibly small compatibility helpers in `utils.py` imports.
- Affected tests: new unit tests for workflow/service behavior plus updates to CLI and API integration tests.
- Affected docs/specs: new `workflow-service` capability and delta specs for affected entry-point capabilities.
- No new runtime dependency is expected.
- Matching algorithm changes, `utils.py` module split, and API envelope versioning are out of scope for this change.
