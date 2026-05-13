## MODIFIED Requirements

### Requirement: Workflow service entry points

系统 MUST 提供一个 workflow/service layer，用于暴露 matching-only、mark-only、full sales-report 和 API-oriented merge workflows 的 application-level operations，同时通过稳定的 compatibility functions 将 business rules 委派给重构后的 core modules。

说明：service 层是“编排边界”，不是新的业务算法层。它可以决定调用哪个 core workflow、何时写回文件、如何汇总统计和归一化错误；但不应内联 exact/P-number/hyphen 匹配或 CSV fallback 细节。

#### Scenario: Match-only workflow service
- **WHEN** 某个 entry point 使用 order file 和 payment file 调用 match-only service operation 时
- **THEN** service SHALL 运行现有 payment-fee matching logic
- **AND** service SHALL 返回包含 output file path、updated DataFrame 和 match statistics 的 structured result
- **AND** service SHALL NOT 在内联中重复 core matching algorithm

#### Scenario: Mark-only workflow service
- **WHEN** 某个 entry point 使用 order file 调用 mark-only service operation 时
- **THEN** service SHALL 运行现有 sales-report period marking logic
- **AND** service SHALL 返回包含 output file path、updated DataFrame 和 marking statistics 的 structured result

#### Scenario: Full sales-report workflow service
- **WHEN** 某个 entry point 使用 order file、payment file 和 `target_month` 调用 sales-report service operation 时
- **THEN** service SHALL 运行现有 full sales-report workflow
- **AND** service SHALL 返回包含 output file path、updated order DataFrame、report DataFrame 和 full workflow statistics 的 structured result
- **AND** service SHALL 保留现有 no-CLI-report-file contract

#### Scenario: API-oriented workflow service
- **WHEN** HTTP API 使用 uploaded file paths 和可选 `month` 调用 API-oriented workflow operation 时
- **THEN** service SHALL 产出 API routes 所需的 result path、download name、download URL、statistics 和 file metadata
- **AND** service SHALL 保留 API-specific downloadable artifact behavior

### Requirement: Centralized statistics calculation

workflow/service layer MUST 为共享 workflows 集中统计计算，避免 entry points 重复公式。重构 core modules SHALL NOT 将 adapter output formatting 移入 statistics layer。

说明：集中的是统计公式和错误归一化，不是传输层格式。CLI 的 `ok/data/error` envelope 和 API 的既有响应 shape 仍分别由各自 adapter 负责。

#### Scenario: Match statistics
- **WHEN** 某个 matching result DataFrame 被处理时
- **THEN** service SHALL 计算 `total_rows`、`matched_rows` 和 `match_rate`

#### Scenario: Mark statistics
- **WHEN** 某个 marked order DataFrame 被处理时
- **THEN** service SHALL 计算 `total_rows` 和 `marked_rows`

#### Scenario: Full workflow statistics
- **WHEN** 某个 full sales-report workflow result 被处理时
- **THEN** service SHALL 计算 `total_rows`、`matched_rows`、`match_rate` 和 `marked_rows`

#### Scenario: API report statistics
- **WHEN** 某个 API sales-report request 产出 filtered report DataFrame 时
- **THEN** service SHALL 在 API-facing statistics 中包含 full workflow statistics 以及 `report_rows`

#### Scenario: Empty API report data
- **WHEN** 在当前 contract 下，API sales-report request 产出空的 filtered report DataFrame 时
- **THEN** service SHALL 抛出 `code="processing_error"` 的 `WorkflowError`
- **AND** 它 SHALL NOT 返回 downloadable report metadata

#### Scenario: Adapters consume service statistics
- **WHEN** CLI、interactive mode 或 HTTP API 需要 workflow statistics 时
- **THEN** 它 SHALL 使用 service result 或 service error mapping 中的 statistics
- **AND** 它 SHALL NOT 独立重新计算共享公式

## ADDED Requirements

### Requirement: Adapter and service boundary preservation

workflow/service layer SHALL 继续作为 adapters 与 core modules 之间的 application orchestration boundary。CLI 和 HTTP output formatting MUST 保持在 adapters 中。

#### Scenario: CLI adapter owns CLI transport formatting
- **WHEN** 某个 CLI request 成功完成或失败时
- **THEN** `cli.py` SHALL 继续负责 stdout/stderr output、JSON envelope formatting 和 process exit codes
- **AND** `workflow_service.py` SHALL 提供 structured results 或 normalized errors，而不是打印 CLI responses

#### Scenario: HTTP adapter owns HTTP transport formatting
- **WHEN** 某个 API request 成功完成或失败时
- **THEN** `excel_merge_api.py` SHALL 继续负责 Flask response objects、HTTP status codes 和 API-specific JSON shape
- **AND** `workflow_service.py` SHALL NOT 输出 Flask response objects

#### Scenario: Core modules are invoked through service operations
- **WHEN** CLI、interactive 或 API entry points 需要 matching 或 sales-report workflows 时
- **THEN** they SHOULD 调用 workflow/service operations，而不是直接协调多个 core functions
- **AND** 任何直接使用 core function 的行为 SHALL 保留已文档化的 adapter output contracts
