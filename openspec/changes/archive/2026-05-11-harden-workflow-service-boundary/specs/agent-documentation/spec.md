## MODIFIED Requirements

### Requirement: AGENTS.md CLI 使用参考章节

AGENTS.md MUST contain a dedicated CLI usage reference that reflects the current `cli.py` contract: two required positional files, optional positional `target_month`, explicit reduced mode flags, JSON/text output, and exit codes. The documentation MUST NOT describe `--month` as a currently supported `cli.py` parameter unless clearly marked as a future or alternate interface. The documentation MUST state that CLI merge and sales-report results are written in place to the order file.

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
- **AND** 章节中不将 `--month`、`-o`、`--output`、`--output-dir` 描述为当前可用参数

#### Scenario: 默认完整工作流文档
- **WHEN** Agent 需要了解默认工作流
- **THEN** AGENTS.md 中包含完整销售报表工作流的命令示例 using positional `target_month`
- **AND** 示例展示提供 `target_month` 后就地修改订单文件
- **AND** 说明如果缺少月份，Agent/Skill 应先从文件名或上下文推断，无法推断时询问用户
- **AND** 说明不应在缺少月份时静默降级为基础匹配
- **AND** 不包含任何指定独立输出文件的命令示例
- **AND** 说明文件格式支持 Excel (.xlsx, .xls) 和 CSV

### Requirement: 文档一致性

项目中所有面向 AI Agent 和用户的 CLI 文档 MUST 保持一致，反映 cli.py 的实际行为。被移除的参数与 CLI 产物 MUST NOT 被描述为可用能力；文档 MAY 在错误场景或迁移说明中提及被移除参数。

#### Scenario: AGENTS.md 与 USAGE_EXAMPLES.md 一致性
- **WHEN** 对比 AGENTS.md 和 documents/USAGE_EXAMPLES.md 中的 CLI 信息
- **THEN** 参数列表、退出码表、JSON 格式在语义上一致
- **AND** 中英文说明准确对应（USAGE_EXAMPLES.md 使用中文）
- **AND** 两份文档均不演示 `--month`、`-o`、`--output-dir` 作为当前可用 CLI 参数或 `report_*.xlsx` 作为 CLI 产物
- **AND** 命令示例覆盖 full workflow with positional `target_month` and explicit `--match-only` reduced workflow

#### Scenario: SKILL 文档与 CLI 实现一致性
- **WHEN** 对比 `.opencode/skills/excel-merge-cli/SKILL.md` 与 cli.py 实现
- **THEN** SKILL 中描述的参数集合等于 cli.py argparse 注册的集合
- **AND** SKILL 不将 `--month`、`-o`、`--output`、`--output-dir` 描述为当前可用参数
- **AND** SKILL 不将 `report_YYYYMM.xlsx` 描述为 CLI 产物或其下载/落地路径
- **AND** SKILL 的 JSON shape 示例与 `cli-output` capability 中定义的一致
