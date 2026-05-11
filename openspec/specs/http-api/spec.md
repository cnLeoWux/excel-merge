## Purpose

HTTP API 能力 - 定义 `excel_merge_api.py` 提供的 Flask HTTP 接口契约，包括端点路径、请求格式、响应格式与文件上传/下载流程。该能力面向需要通过 HTTP 集成本工具的内部服务或 Web 前端。

## Requirements

### Requirement: 端点清单

Flask 应用 MUST 提供以下端点，路径与方法不得变更（向后兼容）。

#### Scenario: 上传页面
- **WHEN** 客户端发送 `GET /`
- **THEN** 返回 HTML 上传表单（200）
- **AND** 表单支持选择订单文件和支付文件并提交到 `/merge`

#### Scenario: 健康检查
- **WHEN** 客户端发送 `GET /health`
- **THEN** 返回 JSON `{"status": "ok"}`（或等价结构），HTTP 200

#### Scenario: 合并并返回文件
- **WHEN** 客户端发送 `POST /merge`，附带 `order_file` 与 `payment_file` 两个 multipart 字段
- **THEN** 处理完成后返回处理后的文件作为 attachment 下载

#### Scenario: 合并并返回 JSON
- **WHEN** 客户端发送 `POST /merge/json`，附带 `order_file` 与 `payment_file`
- **THEN** 返回 JSON，包含处理统计和结果文件下载 URL

#### Scenario: 下载结果文件
- **WHEN** 客户端发送 `GET /download/<filename>`
- **AND** `results/<filename>` 存在
- **THEN** 返回该文件作为 attachment

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

### Requirement: 文件存储隔离

上传文件与结果文件 MUST 存储到独立目录，并使用安全的文件名以防止路径穿越。

#### Scenario: 上传文件目录
- **WHEN** API 接收到上传文件
- **THEN** 文件保存到 `uploads/` 目录
- **AND** 文件名通过 `werkzeug.utils.secure_filename()` 处理

#### Scenario: 结果文件目录
- **WHEN** 处理生成的结果文件
- **THEN** 文件保存到 `results/` 目录
- **AND** 文件名通过 `secure_filename()` 处理

#### Scenario: 路径穿越防护
- **WHEN** 客户端请求 `GET /download/../etc/passwd`
- **THEN** 请求被拒绝（404 或 403）
- **AND** 不会读取 `results/` 目录之外的文件

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

### Requirement: 字符编码与 MIME 类型

API SHALL return downloadable files as attachments. `/merge` currently uses the Excel OpenXML MIME type for all returned attachments, while `/download/<filename>` delegates MIME detection to Flask `send_file()`.

#### Scenario: Excel 文件下载 MIME 类型
- **WHEN** 下载 `.xlsx` 文件
- **THEN** 响应 `Content-Type` 为 `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

#### Scenario: CSV 文件下载 MIME 类型
- **WHEN** 下载 `.csv` 文件
- **THEN** `/download/<filename>` SHOULD allow Flask to infer a CSV-compatible MIME type
- **AND** `/merge` MAY still return the Excel OpenXML MIME type for backward compatibility

#### Scenario: JSON 响应字符集
- **WHEN** 任何 JSON 端点返回中文错误消息
- **THEN** 响应正确编码为 UTF-8
- **AND** 客户端能完整解析中文内容

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

### Requirement: HTTP API routes use workflow service

HTTP API merge routes MUST use the workflow/service layer for shared processing while preserving API request validation, response shape, and downloadable artifact behavior. The routes SHALL consume service-produced metadata instead of recomputing shared workflow statistics.

#### Scenario: /merge without month uses API service
- **WHEN** `/merge` receives valid uploaded order and payment files without `month`
- **THEN** the route SHALL save uploads safely
- **AND** it SHALL call the API-oriented matching workflow service operation
- **AND** it SHALL return the service result file as an attachment

#### Scenario: /merge with month uses API service
- **WHEN** `/merge` receives valid uploaded order and payment files with `month`
- **THEN** the route SHALL call the API-oriented sales-report workflow service operation
- **AND** it SHALL return the generated report attachment when report data is produced

#### Scenario: /merge/json without month uses API service
- **WHEN** `/merge/json` receives valid uploaded order and payment files without `month`
- **THEN** the route SHALL call the API-oriented matching workflow service operation
- **AND** it SHALL format the service result using the documented API-specific `success` response shape

#### Scenario: /merge/json with month uses API service
- **WHEN** `/merge/json` receives valid uploaded order and payment files with `month`
- **THEN** the route SHALL call the API-oriented sales-report workflow service operation
- **AND** it SHALL format the service result using the documented API-specific response shape including `statistics.report_rows`

#### Scenario: Invalid month service error
- **WHEN** `/merge` or `/merge/json` receives valid uploaded order and payment files with an invalid `month`
- **THEN** the workflow service SHALL raise `WorkflowError(code="usage_error")`
- **AND** the route SHALL return HTTP 400
- **AND** `/merge/json` SHALL return a JSON failure response with `success=false` and `error`

### Requirement: HTTP adapter remains responsible for HTTP concerns

`excel_merge_api.py` MUST remain responsible for Flask-specific concerns such as request parsing, upload field validation, `secure_filename()`, HTTP status codes, and `send_file()` responses.

#### Scenario: Upload validation before service call
- **WHEN** an API request is missing required files or has invalid filenames
- **THEN** the API route SHALL return an HTTP error before calling the workflow service

#### Scenario: HTTP response formatting after service call
- **WHEN** the workflow service returns a successful API result
- **THEN** the API route SHALL format that result as either a file attachment or API-specific JSON response

#### Scenario: HTTP formatting of service usage error
- **WHEN** the workflow service raises `WorkflowError(code="usage_error")`
- **THEN** the API route SHALL map it to HTTP 400

#### Scenario: HTTP formatting of service file-not-found error
- **WHEN** the workflow service raises `WorkflowError(code="file_not_found")`
- **THEN** the API route SHALL map it to HTTP 404

#### Scenario: HTTP formatting of service processing error
- **WHEN** the workflow service raises `WorkflowError(code="processing_error")`
- **THEN** the API route SHALL map it to HTTP 500
