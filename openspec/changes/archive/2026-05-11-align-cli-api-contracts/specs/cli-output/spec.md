## MODIFIED Requirements

### Requirement: JSON 结构化输出

CLI 在指定 `--json` 标志时 SHALL 将所有结果以 JSON 格式输出到 stdout。JSON 输出 MUST 使用统一信封格式，包含 `ok`（布尔值）、`data`（成功时的数据对象）和 `error`（失败时的错误对象）三个顶层字段。成功时 `data` 通常包含 `output_file`（字符串路径，等于订单文件本身）和 `statistics`。基础匹配与完整销售报表工作流的 `statistics` MUST 包含 `total_rows`、`matched_rows`、`match_rate`；完整销售报表工作流还 MUST 包含 `marked_rows`。`--mark-only` 模式的 `statistics` MUST 包含 `total_rows` 与 `marked_rows`，且 MAY omit `matched_rows` 与 `match_rate`。`data` MUST NOT 包含 `report_file`、`report_rows` 或 `warnings` 字段。

#### Scenario: 基本合并成功时的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --json`
- **AND** 两个文件均存在且格式正确
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 包含 `output_file`（字符串，值等于 `order.xlsx`）和 `statistics`（对象，含 `total_rows`、`matched_rows`、`match_rate`）
- **AND** `data` 不包含 `report_file`、`report_rows`、`warnings` 字段
- **AND** `error` 为 `null`
- **AND** 进程退出码为 0

#### Scenario: 文件不存在时的 JSON 错误输出
- **WHEN** 用户执行 `python cli.py nonexistent.xlsx payment.xlsx 202602 --match-only --json`
- **AND** `nonexistent.xlsx` 不存在
- **THEN** stdout 输出有效 JSON，`ok` 为 `false`
- **AND** `error` 包含 `code`（值为 `"file_not_found"`）和 `message`（包含文件名的错误描述）
- **AND** `data` 为 `null`
- **AND** 进程退出码为 3

#### Scenario: 处理异常时的 JSON 错误输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --json`
- **AND** 处理过程中发生异常（包括订单文件写入失败）
- **THEN** stdout 输出有效 JSON，`ok` 为 `false`
- **AND** `error` 包含 `code`（值为 `"processing_error"`）和 `message`
- **AND** 进程退出码为 4

#### Scenario: 销售报表工作流的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --json`
- **AND** 处理成功
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 包含 `output_file` 与 `statistics`
- **AND** `statistics` 包含 `total_rows`、`matched_rows`、`match_rate`、`marked_rows`
- **AND** `data` 不包含 `report_file`、`report_rows`、`warnings` 字段
- **AND** `output_file` 等于订单文件路径（账期标记已就地写回）

#### Scenario: mark-only 模式的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --mark-only --json`
- **AND** 处理成功
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data.output_file` 等于订单文件路径
- **AND** `data.statistics` 包含 `total_rows` 与 `marked_rows`

#### Scenario: 交互取消时的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json`
- **AND** CLI 在读取交互式 target_month 时收到 EOF 或用户取消
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 包含 `message`
- **AND** 进程退出码为 0

### Requirement: stdout/stderr 分离

CLI MUST ensure JSON mode writes machine-readable JSON to stdout without interleaved log text. In text mode, the current implementation MAY write progress information and the final in-place update summary to stdout; it MUST NOT print independent "Result saved to:" or "Report saved to:" paths because CLI processing writes back to the order file and does not create an independent report file.

#### Scenario: JSON 模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --json 2>/dev/null`
- **THEN** stdout 仅包含有效 JSON（无日志文本混入）
- **AND** 可通过 `json.loads()` 正确解析

#### Scenario: 文本模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only 2>/dev/null`
- **THEN** stdout 包含处理进度和表示订单文件已就地更新的摘要行（包含 `order.xlsx` 的路径）
- **AND** stdout 不出现指向其它路径的"另存为"信息

#### Scenario: 销售报表工作流文本模式输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602`
- **AND** 处理成功
- **THEN** stdout 输出仅指向 `order.xlsx` 的就地更新摘要
- **AND** stdout 不出现指向 `report_*.xlsx` 的路径
- **AND** 当前工作目录与任何其它目录均不存在新生成的 `report_*.xlsx` 文件
