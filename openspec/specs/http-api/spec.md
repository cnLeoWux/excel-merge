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
- **WHEN** 上传文件扩展名不在白名单内
- **THEN** 返回 HTTP 4xx 错误
- **AND** 错误消息说明支持的格式

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

`/merge/json` 端点 MUST 返回与 CLI `--json` 输出语义一致的 JSON 信封。

#### Scenario: 成功响应
- **WHEN** `/merge/json` 处理成功
- **THEN** 响应 JSON 包含 `ok=true`、`data`（含 `output_file`、`download_url`、`statistics`）、`error=null`
- **AND** `download_url` 指向 `/download/<filename>`

#### Scenario: 失败响应
- **WHEN** `/merge/json` 处理失败
- **THEN** 响应 JSON 包含 `ok=false`、`data=null`、`error`（含 `code` 与 `message`）
- **AND** HTTP 状态码反映错误类型（4xx 客户端错误，5xx 服务端错误）

### Requirement: 字符编码与 MIME 类型

API MUST 在响应中正确声明字符编码与 MIME 类型，避免下载文件被错误识别。

#### Scenario: Excel 文件下载 MIME 类型
- **WHEN** 下载 `.xlsx` 文件
- **THEN** 响应 `Content-Type` 为 `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

#### Scenario: CSV 文件下载 MIME 类型
- **WHEN** 下载 `.csv` 文件
- **THEN** 响应 `Content-Type` 为 `text/csv; charset=utf-8`

#### Scenario: JSON 响应字符集
- **WHEN** 任何 JSON 端点返回中文错误消息
- **THEN** 响应正确编码为 UTF-8
- **AND** 客户端能完整解析中文内容

### Requirement: API sales report trigger
The Flask API endpoints `/merge` and `/merge/json` MUST support triggering the sales report workflow via a form parameter.

#### Scenario: Trigger sales report via /merge/json
- **WHEN** a client sends a `POST` request to `/merge/json` with valid `order_file`, `payment_file`, and a `month` form parameter (e.g., "202602")
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the JSON response `data` field MUST include `report_file` and `report_rows` keys, consistent with the CLI's JSON output for the sales report workflow.

#### Scenario: Trigger sales report via /merge
- **WHEN** a client sends a `POST` request to `/merge` with valid `order_file`, `payment_file`, and a `month` form parameter.
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the system SHALL return the generated monthly report file (`report_YYYYMM.xlsx`) as a file attachment if it was created.

#### Scenario: No month parameter provided
- **WHEN** a client sends a `POST` request to `/merge` or `/merge/json` without the `month` parameter.
- **THEN** the standard matching workflow SHALL be executed.
- **AND** the response SHALL NOT contain sales report artifacts.
