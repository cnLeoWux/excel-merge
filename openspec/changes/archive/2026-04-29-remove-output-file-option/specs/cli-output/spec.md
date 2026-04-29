## MODIFIED Requirements

### Requirement: JSON 结构化输出

CLI 在指定 `--json` 标志时 SHALL 将所有结果以 JSON 格式输出到 stdout。JSON 输出 MUST 使用统一信封格式，包含 `ok`（布尔值）、`data`（成功时的数据对象）和 `error`（失败时的错误对象）三个顶层字段。`data` 的形状 MUST 与是否传入 `--month` 无关：始终仅包含 `output_file`（字符串路径，等于订单文件本身）和 `statistics`（对象，含 `total_rows`、`matched_rows`、`match_rate`）。`data` MUST NOT 包含 `report_file`、`report_rows` 或 `warnings` 字段。


#### Scenario: 基本合并成功时的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json`
- **AND** 两个文件均存在且格式正确
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 包含 `output_file`（字符串，值等于 `order.xlsx`）和 `statistics`（对象，含 `total_rows`、`matched_rows`、`match_rate`）
- **AND** `data` 不包含 `report_file`、`report_rows`、`warnings` 字段
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
- **AND** 处理过程中发生异常（包括订单文件写入失败）
- **THEN** stdout 输出有效 JSON，`ok` 为 `false`
- **AND** `error` 包含 `code`（值为 `"processing_error"`）和 `message`
- **AND** 进程退出码为 4

#### Scenario: 销售报表工作流的 JSON 输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --month 202602 --json`
- **AND** 处理成功
- **THEN** stdout 输出有效 JSON，`ok` 为 `true`
- **AND** `data` 形状与不传 `--month` 时完全一致：仅含 `output_file` 与 `statistics`
- **AND** `data` 不包含 `report_file`、`report_rows`、`warnings` 字段
- **AND** `output_file` 等于订单文件路径（账期标记已就地写回）

### Requirement: stdout/stderr 分离

CLI MUST 将数据输出和日志输出分离：数据（JSON 或最终结果路径）输出到 stdout，日志/进度/警告输出到 stderr。文本模式下，stdout 上仅出现"订单文件就地更新"类的最终摘要，不得出现独立的"Result saved to:"或"Report saved to:"提示，因为不再产生独立的结果文件或月报文件。

#### Scenario: JSON 模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json 2>/dev/null`
- **THEN** stdout 仅包含有效 JSON（无日志文本混入）
- **AND** 可通过 `json.loads()` 正确解析

#### Scenario: 文本模式下的输出分离
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 2>/dev/null`
- **THEN** stdout 输出表示订单文件已就地更新的摘要行（包含 `order.xlsx` 的路径）
- **AND** stdout 不出现指向其它路径的"另存为"信息
- **AND** stderr 包含处理进度日志（如果未使用 `--quiet`）

#### Scenario: 销售报表工作流文本模式输出
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --month 202602`
- **AND** 处理成功
- **THEN** stdout 输出仅指向 `order.xlsx` 的就地更新摘要
- **AND** stdout 不出现指向 `report_*.xlsx` 的路径
- **AND** 当前工作目录与任何其它目录均不存在新生成的 `report_*.xlsx` 文件

## REMOVED Requirements

### Requirement: 工作流 JSON 输出扩展

**Reason**: 销售报表工作流不再产出独立报表文件，所有结果（含 `销售报表账期` 列）就地写回订单文件，因此 JSON 信封不再需要 `report_file` / `report_rows` 等月报相关字段。"部分成功 + warnings" 路径同时被取消：写订单文件失败一律以 `processing_error`（退出码 4）失败。

**Migration**: 调用方不再读取 `data.report_file`、`data.report_rows`、`data.warnings`。若历史代码依赖这些字段，应改为：(1) 检查 `ok` 与退出码判断成功；(2) 通过 `data.statistics.matched_rows` 评估匹配规模；(3) 写入失败现在直接以非零退出码 + `error.code == "processing_error"` 暴露，不再有"成功但有 warning"的中间态。
