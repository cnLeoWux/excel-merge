## Purpose

CLI 输出能力 - 定义命令行工具的结构化输出、退出码、日志控制等行为规范，支持 AI Agent 和自动化脚本集成。

## Requirements

### Requirement: JSON 结构化输出

CLI 在指定 `--json` 标志时 SHALL 将所有结果以 JSON 格式输出到 stdout。JSON 输出 MUST 使用统一信封格式，包含 `ok`（布尔值）、`data`（成功时的数据对象）和 `error`（失败时的错误对象）三个顶层字段。成功时 `data` 通常包含 `output_file`（字符串路径，等于订单文件本身）和 `statistics`，且这些字段 SHALL 来自 workflow/service 层的结果对象。基础匹配与完整销售报表工作流的 `statistics` MUST 包含 `total_rows`、`matched_rows`、`match_rate`；完整销售报表工作流还 MUST 包含 `marked_rows`。`--mark-only` 模式的 `statistics` MUST 包含 `total_rows` 与 `marked_rows`，且 MAY omit `matched_rows` 与 `match_rate`。`data` MUST NOT 包含 `report_file`、`report_rows` 或 `warnings` 字段。


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

### Requirement: 语义化退出码

CLI MUST 在所有退出路径上使用语义化退出码，以便调用方（AI Agent 或脚本）根据退出码判断错误类型。

#### Scenario: 成功退出
- **WHEN** CLI 处理完成且无错误
- **THEN** 进程退出码为 0

#### Scenario: 参数错误退出
- **WHEN** CLI 接收到无效参数（如缺少必需的位置参数）
- **THEN** 进程退出码为 2（argparse 默认行为）

#### Scenario: 文件未找到退出
- **WHEN** 指定的输入文件不存在
- **THEN** 进程退出码为 3

#### Scenario: 处理错误退出
- **WHEN** 文件读取或匹配过程中发生异常
- **THEN** 进程退出码为 4

#### Scenario: 通用错误退出
- **WHEN** 发生未预期的异常
- **THEN** 进程退出码为 1

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

### Requirement: 日志级别控制

CLI MUST 支持 `--quiet` 和 `--verbose` 标志以控制日志详细程度。

#### Scenario: 静默模式
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --quiet`
- **THEN** stderr 上不输出进度日志（仅输出警告和错误）
- **AND** 处理结果正常输出到 stdout

#### Scenario: 详细模式
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --verbose`
- **THEN** stderr 上输出详细的匹配过程日志（包括每行的匹配尝试）

#### Scenario: 默认模式（无标志）
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only`（不带 `--quiet` 或 `--verbose`）
- **THEN** 行为与当前版本一致（输出处理进度摘要并写回订单文件）
- **AND** 向后兼容

### Requirement: 处理前自动备份

CLI MUST create a timestamped backup of the order file before processing any mode that may write results back to the order file.

#### Scenario: Backup before processing
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --json --quiet`
- **AND** 两个输入文件存在
- **THEN** CLI 在处理前调用自动备份逻辑
- **AND** 备份文件被写入 `backup/` 目录

#### Scenario: Backup failure
- **WHEN** 自动备份失败
- **THEN** CLI 将错误作为处理错误输出
- **AND** JSON 模式下 `error.code` 为 `"processing_error"`
- **AND** 进程退出码为 4

### Requirement: CLI output uses workflow service results

CLI output formatting MUST consume workflow/service result objects for `data.output_file`, `data.statistics`, and processing errors instead of recomputing shared workflow statistics in `cli.py`.

#### Scenario: Full workflow JSON from service result
- **WHEN** `cli.py` completes a full sales-report workflow through the service layer
- **THEN** CLI JSON `data.output_file` SHALL come from the service result
- **AND** CLI JSON `data.statistics` SHALL come from the service result

#### Scenario: Reduced workflow JSON from service result
- **WHEN** `cli.py` completes `--match-only` or `--mark-only` through the service layer
- **THEN** CLI JSON statistics SHALL reflect the service result for that selected mode

#### Scenario: Error mapping from service error
- **WHEN** the service layer returns or raises a normalized workflow error
- **THEN** `cli.py` SHALL map it to the documented CLI JSON error envelope and exit code

### Requirement: CLI adapter remains responsible for transport formatting

`cli.py` MUST remain responsible for argument parsing, interactive prompting, stdout/stderr formatting, and `sys.exit()` behavior even after workflow execution moves into the service layer.

#### Scenario: Argument parsing remains in CLI
- **WHEN** a user invokes `cli.py`
- **THEN** `cli.py` SHALL parse positional file arguments, optional `target_month`, mode flags, and output flags before calling the service layer

#### Scenario: JSON envelope remains in CLI
- **WHEN** CLI output is emitted in JSON mode
- **THEN** `cli.py` SHALL format the service result using the documented `ok/data/error` envelope
