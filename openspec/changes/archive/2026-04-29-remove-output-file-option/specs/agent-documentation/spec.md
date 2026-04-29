## MODIFIED Requirements

### Requirement: AGENTS.md CLI 使用参考章节

AGENTS.md MUST 包含一个专门的 CLI 使用参考章节，完整覆盖 cli.py 的全部参数、工作流、输出格式和退出码。文档 MUST NOT 描述、示例化或暗示已被移除的 `-o`/`--output`、`--output-dir` 参数，亦 MUST NOT 提及 `report_YYYYMM.xlsx` 等独立报表文件产物。文档 MUST 明确说明：所有合并与销售报表标记结果一律就地写回订单文件。

#### Scenario: CLI 参数完整性
- **WHEN** Agent 读取 AGENTS.md 的 CLI 使用参考章节
- **THEN** 章节中包含所有 CLI 参数的说明表格
- **AND** 每个参数包含名称、类型、默认值和说明
- **AND** 仅覆盖以下参数（不多不少）：
  - `order_file`（必填位置参数）
  - `payment_file`（必填位置参数）
  - `--month`（可选，触发销售报表工作流）
  - `--json`（可选，JSON 输出模式）
  - `--quiet`（可选，静默模式）
  - `-v`, `--verbose`（可选，详细日志模式）
- **AND** 章节中不出现 `-o`、`--output`、`--output-dir` 任一字符串

#### Scenario: 参数默认值文档准确性
- **WHEN** Agent 读取 AGENTS.md 中某参数的默认值说明
- **THEN** 该说明与 cli.py 中 argparse 定义的 `default` 值一致
- **AND** `--month` 的默认值说明为"无"或"None"
- **AND** `--json` 的默认值说明为"False"或"不启用"

#### Scenario: 基本工作流文档
- **WHEN** Agent 需要了解基本匹配工作流
- **THEN** AGENTS.md 中包含基本工作流的命令示例
- **AND** 示例展示默认用法（就地修改订单文件）
- **AND** 不包含任何指定独立输出文件的命令示例
- **AND** 说明文件格式支持 Excel (.xlsx, .xls) 和 CSV

#### Scenario: JSON 输出格式文档
- **WHEN** Agent 需要解析 CLI 的 JSON 输出
- **THEN** AGENTS.md 中包含成功和失败两种 JSON 信封格式的示例
- **AND** 成功信封 `data` 仅包含以下字段：
  - `output_file`: string（值等于订单文件路径）
  - `statistics`: { `total_rows`: number, `matched_rows`: number, `match_rate`: string }
- **AND** 成功信封 `data` 不出现 `report_file`、`report_rows`、`warnings` 字段
- **AND** 失败信封包含以下字段：
  - `ok`: false
  - `data`: null
  - `error`: { `code`: string, `message`: string }
- **AND** 错误代码包含：`file_not_found`, `processing_error`, `unknown_error`

#### Scenario: 退出码文档完整性
- **WHEN** Agent 需要判断 CLI 执行结果
- **THEN** AGENTS.md 中包含完整的退出码表
- **AND** 覆盖以下退出码：
  - 0: 成功 (EXIT_SUCCESS)
  - 1: 通用错误 (EXIT_GENERAL_ERROR)
  - 2: 用法错误 (EXIT_USAGE_ERROR)
  - 3: 文件未找到 (EXIT_FILE_NOT_FOUND)
  - 4: 处理错误 (EXIT_PROCESSING_ERROR)
- **AND** 每个退出码包含语义说明和典型触发场景
- **AND** 退出码 2 的典型场景明确包含"传入已被移除的 `-o`/`--output-dir` 等无效参数"

#### Scenario: 退出码与代码一致性
- **WHEN** 对比 AGENTS.md 中的退出码定义与 cli.py 中的常量
- **THEN** 退出码数值与 cli.py 中定义的常量一致
- **AND** 退出码语义说明与 cli.py 中 sys.exit() 调用场景匹配

#### Scenario: Agent 推荐调用方式文档
- **WHEN** Agent 需要以非交互方式调用本工具
- **THEN** AGENTS.md 中包含推荐的 Agent 调用命令示例
- **AND** 推荐使用 `--json --quiet` 组合
- **AND** 说明 `--json` 将结果输出到 stdout
- **AND** 说明 `--quiet` 抑制进度日志到 stderr
- **AND** 说明 stdout 仅输出 JSON、stderr 输出日志的分离规则

#### Scenario: stdout/stderr 分离规则文档
- **WHEN** Agent 需要理解 CLI 的输出流设计
- **THEN** AGENTS.md 中明确说明：
  - stdout: 仅输出 JSON 结果（`--json` 模式）或就地更新摘要（文本模式）
  - stderr: 输出日志、进度信息、警告和错误消息
- **AND** 说明如何通过 `capture_output=True` 分别获取 stdout 和 stderr

#### Scenario: 销售报表工作流文档
- **WHEN** Agent 需要了解 `--month` 触发的销售报表工作流
- **THEN** AGENTS.md 中包含销售报表工作流的完整说明
- **AND** 说明 `--month YYYYMM` 参数格式和作用
- **AND** 不出现 `--output-dir` 参数或对其的引用
- **AND** 说明两阶段处理流程：
  1. 匹配支付手续费
  2. 标记 `销售报表账期` 列（"全退"和"已取消"）
  3. 在内存中筛选未标记且出行日期在目标月份前后 1 年范围内的数据
  4. 将更新后的订单 DataFrame 就地写回原始订单文件
- **AND** 明确说明工作流不产生 `report_YYYYMM.xlsx` 等独立报表文件
- **AND** 包含完整的命令示例

#### Scenario: 常见错误场景文档
- **WHEN** Agent 需要帮助用户调试 CLI 错误
- **THEN** AGENTS.md 中包含常见错误场景及处理方式
- **AND** 覆盖以下场景：
  - 文件不存在 (退出码 3, error.code="file_not_found")
  - 处理错误 (退出码 4, error.code="processing_error")，包括订单文件写入失败
  - 用法错误 (退出码 2)，包括传入已被移除的 `-o`/`--output-dir`
- **AND** 每个场景包含错误信息示例和解决建议

### Requirement: 文档一致性

项目中所有面向 AI Agent 和用户的 CLI 文档 MUST 保持一致，反映 cli.py 的实际行为。被移除的参数与产物 MUST NOT 出现在任何文档（含 README.md、AGENTS.md、`documents/USAGE_EXAMPLES.md`、`.opencode/skills/excel-merge-cli/SKILL.md`）中。

#### Scenario: AGENTS.md 与 README.md CLI 参数一致性
- **WHEN** 对比 AGENTS.md 和 README.md 中的 CLI 参数列表
- **THEN** 两者列出的参数集合相同
- **AND** 两者均不出现 `-o`、`--output`、`--output-dir`
- **AND** 参数说明不矛盾
- **AND** 参数默认值描述一致

#### Scenario: AGENTS.md 与 README.md 退出码一致性
- **WHEN** 对比 AGENTS.md 和 README.md 中的退出码表
- **THEN** 两者的退出码数值和语义说明一致
- **AND** 覆盖的退出码集合相同 (0, 1, 2, 3, 4)

#### Scenario: AGENTS.md 与 README.md JSON 格式一致性
- **WHEN** 对比 AGENTS.md 和 README.md 中的 JSON 格式示例
- **THEN** 成功信封的字段名和嵌套结构相同
- **AND** 两者的 `data` 均不含 `report_file`、`report_rows`、`warnings`
- **AND** 失败信封的字段名和嵌套结构相同
- **AND** error.code 可能值集合一致

#### Scenario: AGENTS.md 与 USAGE_EXAMPLES.md 一致性
- **WHEN** 对比 AGENTS.md 和 documents/USAGE_EXAMPLES.md 中的 CLI 信息
- **THEN** 参数列表、退出码表、JSON 格式在语义上一致
- **AND** 中英文说明准确对应（USAGE_EXAMPLES.md 使用中文）
- **AND** 两份文档均不演示 `-o`、`--output-dir` 或 `report_*.xlsx` 产物
- **AND** 命令示例覆盖相同的用例场景

#### Scenario: SKILL 文档与 CLI 实现一致性
- **WHEN** 对比 `.opencode/skills/excel-merge-cli/SKILL.md` 与 cli.py 实现
- **THEN** SKILL 中描述的参数集合等于 cli.py argparse 注册的集合
- **AND** SKILL 不出现 `-o`、`--output`、`--output-dir`
- **AND** SKILL 不出现 `report_YYYYMM.xlsx` 产物或其下载/落地路径
- **AND** SKILL 的 JSON shape 示例与 `cli-output` capability 中定义的一致

#### Scenario: 文档与实际代码一致性
- **WHEN** 对比 AGENTS.md 中的 CLI 文档与 cli.py 中的实现
- **THEN** 所有 argparse 参数在文档中有对应说明
- **AND** 默认值描述与代码中 `parser.add_argument(..., default=...)` 一致
- **AND** 退出码常量与文档中的数值一致
- **AND** JSON 输出格式与 `output_result()` 函数输出一致
- **AND** 工作流描述与 `main_cli()` 函数逻辑一致

#### Scenario: 销售报表工作流文档与代码一致性
- **WHEN** 对比文档中的销售报表工作流说明与 cli.py 实现
- **THEN** 描述的处理流程与 `main_cli()` 中 `--month` 分支逻辑一致
- **AND** 文档中明确"工作流不产生独立报表文件"，与代码中无 `to_excel(report_*)` 调用一致
- **AND** 文档中描述的最终产物为"就地更新的订单文件"，与 `write_result_file(updated_df, order_file)` 行为一致
