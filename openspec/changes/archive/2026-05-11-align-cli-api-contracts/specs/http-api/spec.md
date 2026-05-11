## MODIFIED Requirements

### Requirement: 文件上传约束

API MUST 校验上传文件的格式与大小，拒绝不合法请求并返回明确错误。

#### Scenario: 支持的文件格式
- **WHEN** 上传文件扩展名为 `.xlsx`、`.xls` 或 `.csv`
- **THEN** 接受并继续处理

#### Scenario: 不支持的格式
- **WHEN** 客户端向 `/merge` 上传的文件扩展名不在白名单内
- **THEN** 返回 HTTP 4xx 错误
- **AND** 错误消息说明支持的格式

#### Scenario: /merge/json 当前扩展名校验
- **WHEN** 客户端向 `/merge/json` 上传文件
- **THEN** 当前实现 MAY accept files without applying the same `allowed_file()` extension check used by `/merge`
- **AND** downstream processing errors are returned as JSON errors

#### Scenario: 文件大小上限
- **WHEN** 上传文件总大小超过 16MB
- **THEN** 返回 HTTP 413 (Payload Too Large)
- **AND** 错误消息说明大小限制

#### Scenario: 缺少必需文件字段
- **WHEN** 请求缺少 `order_file` 或 `payment_file`
- **THEN** 返回 HTTP 4xx 错误
- **AND** 错误消息指出缺失的字段名

### Requirement: JSON 响应格式

`/merge/json` 端点 SHALL return the current API-specific JSON shape. This endpoint is not currently required to use the CLI `--json` envelope.

#### Scenario: 成功响应
- **WHEN** `/merge/json` 处理成功
- **THEN** 响应 JSON 包含 `success=true`
- **AND** 响应 JSON 包含 `session_id`、`download_url`、`statistics`、`files`
- **AND** `download_url` 指向 `/download/<filename>`
- **AND** 响应 JSON 不需要包含顶层 `ok`、`data` 或 `error` 字段

#### Scenario: 失败响应
- **WHEN** `/merge/json` 处理失败
- **THEN** 响应 JSON 包含 `success=false` 和 `error` 字段，或在请求校验失败时包含 `error` 字段
- **AND** HTTP 状态码反映错误类型（4xx 客户端错误，5xx 服务端错误）

### Requirement: API sales report trigger

The Flask API endpoints `/merge` and `/merge/json` MUST support triggering the sales report workflow via a form parameter.

#### Scenario: Trigger sales report via /merge/json
- **WHEN** a client sends a `POST` request to `/merge/json` with valid `order_file`, `payment_file`, and a `month` form parameter (e.g., "202602")
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the API layer SHALL persist the filtered report DataFrame to a downloadable file under `results/` (the workflow function itself does not write files).
- **AND** the JSON response MUST include `success=true`, a `download_url` pointing to that file, and a `statistics.report_rows` integer count. (This is the API's own response shape and is independent from the CLI JSON envelope, which does not carry `report_file`/`report_rows`.)

#### Scenario: Trigger sales report via /merge
- **WHEN** a client sends a `POST` request to `/merge` with valid `order_file`, `payment_file`, and a `month` form parameter.
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the system SHALL return the generated monthly report file (`report_YYYYMM.xlsx`) as a file attachment if it was created.

#### Scenario: No month parameter provided
- **WHEN** a client sends a `POST` request to `/merge` or `/merge/json` without the `month` parameter.
- **THEN** the standard matching workflow SHALL be executed.
- **AND** the response SHALL NOT contain sales report artifacts.
