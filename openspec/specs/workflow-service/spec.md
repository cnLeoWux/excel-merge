## Purpose

Workflow service 能力 - 定义应用级 workflow/service 层如何协调 `utils.py` 中已有的匹配、销售报表、文件写回、API 结果文件与统计构建逻辑，使 CLI、交互入口和 HTTP API 能共享编排行为而不重复业务流程。

## Requirements

### Requirement: Workflow service 入口

系统 MUST 提供 workflow/service 层，暴露仅匹配、仅标注、完整销售报表以及面向 API 的合并工作流的应用级操作，同时将现有业务规则委托给 `utils.py` 中的函数。

#### Scenario: 仅匹配 workflow service
- **WHEN** 入口点使用订单文件和支付文件调用仅匹配 service 操作
- **THEN** service SHALL 执行现有的支付手续费匹配逻辑
- **AND** service SHALL 返回包含输出文件路径、更新后的 DataFrame 与匹配统计的结构化结果

#### Scenario: 仅标注 workflow service
- **WHEN** 入口点使用订单文件调用仅标注 service 操作
- **THEN** service SHALL 执行现有的销售报表账期标注逻辑
- **AND** service SHALL 返回包含输出文件路径、更新后的 DataFrame 与标注统计的结构化结果

#### Scenario: 完整销售报表 workflow service
- **WHEN** an entry point calls the sales-report service operation with an order file, payment file, and `target_month`
- **THEN** the service SHALL run the existing full sales-report workflow
- **AND** the service SHALL return a structured result containing the output file path, updated order DataFrame, report DataFrame, and full workflow statistics

#### Scenario: 面向 API 的 workflow service
- **WHEN** HTTP API 使用上传文件路径和可选 `month` 调用面向 API 的 workflow 操作
- **THEN** service SHALL 产出路由所需的结果路径、下载名、下载 URL、统计与文件元数据
- **AND** service SHALL 保留 API 专用的可下载产物行为

### Requirement: Workflow 结果结构

workflow/service 层 MUST 返回明确的结构化结果，而不是要求入口点从原始 DataFrame 重新计算统计或推断持久化输出。它还 MUST 在调用核心销售报表工作流前验证目标月份输入。

#### Scenario: CLI 兼容的 workflow 结果
- **WHEN** CLI 或交互式入口点接收到 workflow 结果
- **THEN** 结果 SHALL 包含 `output_file` 与 `statistics`
- **AND** 入口点可将结果格式化为 CLI 文本或 `ok/data/error` JSON 信封，而无需重新计算统计

#### Scenario: API 兼容的 workflow 结果
- **WHEN** API 路由接收到 API workflow 结果
- **THEN** 结果 SHALL 包含 `result_path`、`download_name`、`download_url`、`statistics` 与 `files`
- **AND** 路由可使用现有 API 专用 JSON 结构或文件 attachment 行为格式化响应

#### Scenario: service 级错误结构
- **WHEN** service 层处理已知 workflow 失败
- **THEN** 它 SHALL 暴露规范化的错误码与消息，供入口点映射到 CLI 退出码或 HTTP 状态码

#### Scenario: 缺少输入文件错误
- **WHEN** service 操作接收到不存在的订单文件或支付文件路径
- **THEN** 它 SHALL 抛出 `WorkflowError`，`code="file_not_found"`
- **AND** `exit_code` SHALL 为 3

#### Scenario: 无效目标月份错误
- **WHEN** service 操作接收到非空且不符合 `YYYYMM` 的 `target_month` 或 API `month`
- **THEN** 它 SHALL 抛出 `WorkflowError`，`code="usage_error"`
- **AND** `exit_code` SHALL 为 2
- **AND** 它 SHALL NOT 调用核心销售报表工作流

#### Scenario: 写入失败错误
- **WHEN** service 操作无法回写或持久化其结果文件
- **THEN** 它 SHALL 抛出 `WorkflowError`，`code="processing_error"`
- **AND** `exit_code` SHALL 为 4

### Requirement: 统计计算集中化

workflow/service 层 MUST 为共享工作流集中统计计算，避免入口点重复公式。

#### Scenario: 匹配统计
- **WHEN** a matching result DataFrame is processed
- **THEN** the service SHALL calculate `total_rows`, `matched_rows`, and `match_rate`

#### Scenario: 标注统计
- **WHEN** a marked order DataFrame is processed
- **THEN** the service SHALL calculate `total_rows` and `marked_rows`

#### Scenario: 完整工作流统计
- **WHEN** a full sales-report workflow result is processed
- **THEN** the service SHALL calculate `total_rows`, `matched_rows`, `match_rate`, and `marked_rows`

#### Scenario: API 报表统计
- **WHEN** an API sales-report request produces a filtered report DataFrame
- **THEN** the service SHALL include full workflow statistics plus `report_rows` in API-facing statistics

#### Scenario: 空的 API 报表数据
- **WHEN** an API sales-report request produces an empty filtered report DataFrame under the current contract
- **THEN** the service SHALL raise `WorkflowError` with `code="processing_error"`
- **AND** it SHALL NOT return downloadable report metadata

### Requirement: 持久化协调

workflow/service 层 MUST 根据调用方适配器的契约协调回写或 API 结果文件持久化。

#### Scenario: CLI 就地持久化
- **WHEN** the CLI calls a service operation that writes results
- **THEN** the service SHALL write the updated order DataFrame back to the original order file path
- **AND** the service SHALL NOT create a CLI `report_*.xlsx` artifact

#### Scenario: 交互式就地持久化
- **WHEN** interactive mode calls a service operation that writes results
- **THEN** the service SHALL write the updated order DataFrame back to the selected order file path
- **AND** the service SHALL NOT create an independent report file for interactive mode

#### Scenario: API 可下载持久化
- **WHEN** the API calls a service operation that writes results for download
- **THEN** the service SHALL write the appropriate merged or report DataFrame under the configured API result directory
- **AND** the service SHALL return metadata needed for `/download/<filename>`
