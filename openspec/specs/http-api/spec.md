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

`/merge/json` 端点 SHALL 返回当前 API 专用的 JSON 结构。该端点当前不要求使用 CLI `--json` 信封。

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

API SHALL 将可下载文件作为 attachment 返回。`/merge` 当前对所有返回的 attachment 使用 Excel OpenXML MIME type，而 `/download/<filename>` 则将 MIME 检测委托给 Flask `send_file()`。

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

### Requirement: API 销售报表触发
Flask API 端点 `/merge` 和 `/merge/json` MUST 支持通过表单参数触发销售报表工作流。

#### Scenario: 通过 /merge/json 触发销售报表
- **WHEN** 客户端发送带有有效 `order_file`、`payment_file` 和 `month` 表单参数（例如 "202602"）的 `POST` 请求到 `/merge/json`
- **THEN** `process_sales_report_workflow` SHALL 被执行。
- **AND** API 层 SHALL 将筛选后的 report DataFrame 持久化为 `results/` 下的可下载文件（工作流函数本身不写文件）。
- **AND** JSON 响应 MUST 包含 `success=true`、指向该文件的 `download_url`，以及 `statistics.report_rows` 整数计数。（这是 API 自身的响应结构，与不包含 `report_file`/`report_rows` 的 CLI JSON 信封相互独立。）

#### Scenario: 通过 /merge 触发销售报表
- **WHEN** 客户端发送带有有效 `order_file`、`payment_file` 和 `month` 表单参数的 `POST` 请求到 `/merge`
- **THEN** `process_sales_report_workflow` SHALL 被执行。
- **AND** 如果生成了月度报表文件，系统 SHALL 将其作为文件 attachment 返回（`report_YYYYMM.xlsx`）。

#### Scenario: 未提供月份参数
- **WHEN** 客户端向 `/merge` 或 `/merge/json` 发送未带 `month` 参数的 `POST` 请求
- **THEN** SHALL 执行标准匹配工作流。
- **AND** 响应 SHALL NOT 包含销售报表产物。

### Requirement: HTTP API 路由使用 workflow service

HTTP API merge routes MUST use the workflow/service layer for shared processing while preserving API request validation, response shape, and downloadable artifact behavior. The routes SHALL consume service-produced metadata instead of recomputing shared workflow statistics.

#### Scenario: /merge 无月份使用 API service
- **WHEN** `/merge` 接收到有效的订单文件和支付文件上传，且不带 `month`
- **THEN** 路由 SHALL 安全保存上传文件
- **AND** 它 SHALL 调用面向 API 的匹配 workflow service 操作
- **AND** 它 SHALL 将 service 结果文件作为 attachment 返回

#### Scenario: /merge 带月份使用 API service
- **WHEN** `/merge` 接收到带 `month` 的有效订单文件和支付文件上传
- **THEN** 路由 SHALL 调用面向 API 的销售报表 workflow service 操作
- **AND** 当产出报表数据时，它 SHALL 返回生成的报表 attachment

#### Scenario: /merge/json 无月份使用 API service
- **WHEN** `/merge/json` 接收到有效的订单文件和支付文件上传，且不带 `month`
- **THEN** 路由 SHALL 调用面向 API 的匹配 workflow service 操作
- **AND** 它 SHALL 使用文档化的 API 专用 `success` 响应结构格式化 service 结果

#### Scenario: /merge/json 带月份使用 API service
- **WHEN** `/merge/json` 接收到带 `month` 的有效订单文件和支付文件上传
- **THEN** 路由 SHALL 调用面向 API 的销售报表 workflow service 操作
- **AND** 它 SHALL 使用包含 `statistics.report_rows` 的文档化 API 专用响应结构格式化 service 结果

#### Scenario: 无效月份的 service 错误
- **WHEN** `/merge` 或 `/merge/json` 接收到带无效 `month` 的有效订单文件和支付文件上传
- **THEN** workflow service SHALL 抛出 `WorkflowError(code="usage_error")`
- **AND** 路由 SHALL 返回 HTTP 400
- **AND** `/merge/json` SHALL 返回包含 `success=false` 与 `error` 的 JSON 失败响应

### Requirement: HTTP adapter 仍负责 HTTP 关注点

`excel_merge_api.py` MUST 继续负责 Flask 专属关注点，例如请求解析、上传字段校验、`secure_filename()`、HTTP 状态码与 `send_file()` 响应。

#### Scenario: 调用 service 前的上传校验
- **WHEN** API 请求缺少必需文件或文件名无效
- **THEN** API 路由 SHALL 在调用 workflow service 前返回 HTTP 错误

#### Scenario: service 调用后的 HTTP 响应格式化
- **WHEN** workflow service 返回成功的 API 结果
- **THEN** API 路由 SHALL 将该结果格式化为文件 attachment 或 API 专用 JSON 响应

#### Scenario: service usage error 的 HTTP 格式化
- **WHEN** workflow service 抛出 `WorkflowError(code="usage_error")`
- **THEN** API 路由 SHALL 将其映射为 HTTP 400

#### Scenario: service file-not-found error 的 HTTP 格式化
- **WHEN** workflow service 抛出 `WorkflowError(code="file_not_found")`
- **THEN** API 路由 SHALL 将其映射为 HTTP 404

#### Scenario: service processing error 的 HTTP 格式化
- **WHEN** workflow service 抛出 `WorkflowError(code="processing_error")`
- **THEN** API 路由 SHALL 将其映射为 HTTP 500
