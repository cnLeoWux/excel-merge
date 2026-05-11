## MODIFIED Requirements

### Requirement: AGENTS.md CLI 使用参考章节

AGENTS.md MUST 包含一个专门的 CLI 使用参考章节，完整覆盖 cli.py 的全部参数、工作流、输出格式和退出码。文档 MUST NOT 描述、示例化或暗示已被移除的 `-o`/`--output`、`--output-dir` 参数作为可用 CLI 参数，亦 MUST NOT 将 `report_YYYYMM.xlsx` 描述为 CLI 产物。文档 MUST 明确说明：CLI 的合并与销售报表标记结果一律就地写回订单文件。

#### Scenario: CLI 参数完整性
- **WHEN** Agent 读取 AGENTS.md 的 CLI 使用参考章节
- **THEN** 章节中包含所有 CLI 参数的说明表格
- **AND** 每个参数包含名称、类型、默认值和说明
- **AND** 至少覆盖以下参数：
  - `order_file`（必填位置参数）
  - `payment_file`（必填位置参数）
  - `target_month`（可选位置参数；标准完整流程所需的目标月份）
  - `--match-only`（可选，仅执行匹配，需要 `target_month`）
  - `--mark-only`（可选，仅执行标注，需要 `target_month`）
  - `--json`（可选，JSON 输出模式）
  - `--quiet`（可选，静默模式）
  - `-v`, `--verbose`（可选，详细日志模式）
- **AND** 章节中不将 `-o`、`--output`、`--output-dir` 描述为可用参数

#### Scenario: 默认完整工作流文档
- **WHEN** Agent 需要了解默认工作流
- **THEN** AGENTS.md 中包含完整销售报表工作流的命令示例
- **AND** 示例展示提供 `target_month` 后就地修改订单文件
- **AND** 说明如果缺少月份，Agent/Skill 应先从文件名或上下文推断，无法推断时询问用户
- **AND** 说明不应在缺少月份时静默降级为基础匹配
- **AND** 不包含任何指定独立输出文件的命令示例
- **AND** 说明文件格式支持 Excel (.xlsx, .xls) 和 CSV

#### Scenario: JSON 输出格式文档
- **WHEN** Agent 需要解析 CLI 的 JSON 输出
- **THEN** AGENTS.md 中包含成功和失败两种 JSON 信封格式的示例
- **AND** 成功信封 `data` 通常包含以下字段：
  - `output_file`: string（值等于订单文件路径）
  - `statistics`: { `total_rows`: number, `matched_rows`: number, `match_rate`: string }
- **AND** 完整销售报表工作流的 `statistics` 可额外包含 `marked_rows`
- **AND** `--mark-only` 模式的 `statistics` 可只包含 `total_rows` 与 `marked_rows`
- **AND** 成功信封 `data` 不出现 `report_file`、`report_rows`、`warnings` 字段
- **AND** 失败信封包含以下字段：
  - `ok`: false
  - `data`: null
  - `error`: { `code`: string, `message`: string }

#### Scenario: Agent 缺少月份时的处理
- **WHEN** Agent 已获得订单文件和支付流水文件
- **AND** 用户未明确提供月份
- **THEN** Agent 文档 MUST instruct the Agent to first infer the month from filenames and conversation context
- **AND** if inference is not reliable, ask the user for the target month
- **AND** not run `--match-only` unless the user explicitly requests a reduced matching-only workflow
