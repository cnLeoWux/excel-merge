## ADDED Requirements

### Requirement: JSON 结构化输出

CLI 在指定 `--json` 标志时 SHALL 将所有结果以 JSON 格式输出到 stdout。JSON 输出 MUST 使用统一信封格式，包含 `ok`（布尔值）、`data`（成功时的数据对象）和 `error`（失败时的错误对象）三个顶层字段。

#### Scenario: 基本合并成功时的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json`
- **AND** 两个文件均存在且格式正确
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 包含 `output_file`（字符串）和 `statistics`（对象，含 `total_rows`、`matched_rows`、`match_rate`）
- **AND** `error` 为 `null`
- **AND** 进程退出码为 0

#### Scenario: 文件不存在时的 JSON 错误输出
- **WHEN** 用户执行 `python cli.py nonexistent.xlsx payment.xlsx --json`
- **AND** `nonexistent.xlsx` 不存在
- **THEN** stdout 输出有效 JSON，`ok` 为 `false`
- **AND** `error` 包含 `code`（值为 `"file_not_found"`）和 `message`（包含文件名的错误描述）
- **AND** `data` 为 `null`
- **AND** 进程退出码为 3

#### Scenario: 处理异常时的 JSON 错误输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json`
- **AND** 处理过程中发生异常
- **THEN** stdout 输出有效 JSON，`ok` 为 `false`
- **AND** `error` 包含 `code`（值为 `"processing_error"`）和 `message`
- **AND** 进程退出码为 4

#### Scenario: 销售报表工作流的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --month 202602 --json`
- **AND** 处理成功
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 额外包含 `report_file`（字符串或 `null`）和 `report_rows`（整数或 `null`）

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

CLI MUST 将数据输出和日志输出分离：数据（JSON 或最终结果路径）输出到 stdout，日志/进度/警告输出到 stderr。

#### Scenario: JSON 模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json 2>/dev/null`
- **THEN** stdout 仅包含有效 JSON（无日志文本混入）
- **AND** 可通过 `json.loads()` 正确解析

#### Scenario: 文本模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 2>/dev/null`
- **THEN** stdout 包含结果文件路径信息
- **AND** stderr 包含处理进度日志（如果未使用 `--quiet`）

### Requirement: 日志级别控制

CLI MUST 支持 `--quiet` 和 `--verbose` 标志以控制日志详细程度。

#### Scenario: 静默模式
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --quiet`
- **THEN** stderr 上不输出进度日志（仅输出警告和错误）
- **AND** 处理结果正常输出到 stdout

#### Scenario: 详细模式
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --verbose`
- **THEN** stderr 上输出详细的匹配过程日志（包括每行的匹配尝试）

#### Scenario: 默认模式（无标志）
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx`（不带 `--quiet` 或 `--verbose`）
- **THEN** 行为与当前版本一致（输出处理进度摘要）
- **AND** 向后兼容
