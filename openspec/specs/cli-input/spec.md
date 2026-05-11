## Purpose

CLI 输入能力 - 定义命令行工具的非交互式运行模式，支持 AI Agent 和自动化脚本在无 TTY 环境下调用。

## Requirements

### Requirement: cli.py positional argument mode

`cli.py` MUST expose the current command-line interface using two required positional file arguments and an optional positional `target_month`. The standard workflow is the full sales-report workflow, so callers SHOULD provide or obtain `target_month` unless the user explicitly requests a reduced mode such as `--match-only`.

#### Scenario: Files without target month enter month acquisition
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json --quiet`
- **THEN** `order.xlsx` 被解析为订单文件
- **AND** `payment.xlsx` 被解析为支付流水文件
- **AND** `target_month` 为 `None`
- **AND** CLI 尝试通过交互式提示获取 `target_month`，而不是默认执行基础匹配

#### Scenario: Sales report positional month
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --json --quiet`
- **THEN** `target_month` 被解析为 `202602`
- **AND** CLI 执行完整销售报表工作流

#### Scenario: Missing required positional files
- **WHEN** 用户执行 `python cli.py --json`
- **THEN** argparse 报告缺少必需的位置参数
- **AND** 进程退出码为 2

### Requirement: cli.py mode flags

`cli.py` MUST support `--match-only` and `--mark-only` as mutually exclusive mode flags. In the current contract these flags require the optional positional `target_month` to be present, even though only `--mark-only` uses the month semantically.

#### Scenario: match-only with target month
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --match-only --json`
- **THEN** CLI 仅执行订单匹配并写回订单文件
- **AND** JSON `data.statistics` 包含 `total_rows`、`matched_rows`、`match_rate`

#### Scenario: mark-only with target month
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --mark-only --json`
- **THEN** CLI 仅对订单文件执行销售报表账期标注并写回订单文件
- **AND** JSON `data.statistics` 包含 `total_rows` 与 `marked_rows`

#### Scenario: mode flag without target month
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --match-only --json`
- **THEN** CLI 返回 usage error
- **AND** JSON `error.code` 为 `"usage_error"`
- **AND** 进程退出码为 2

### Requirement: cli.py interactive target_month prompt

When `cli.py` receives only the two file arguments and no mode flag, the current implementation SHALL prompt for `target_month` on stderr before deciding whether to run the full workflow or cancel. This behavior is part of the current contract until a future CLI standardization change replaces it.

#### Scenario: target month entered interactively
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx`
- **AND** 程序提示输入目标月份
- **AND** 用户输入 `202602`
- **THEN** CLI 将 `target_month` 设置为 `202602`
- **AND** 执行完整销售报表工作流

#### Scenario: stdin closed during prompt
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json` 且 stdin 被关闭
- **THEN** CLI 输出成功 JSON 信封
- **AND** `data.message` 表示已取消或无输入
- **AND** 进程退出码为 0

### Requirement: Agent/Skill month acquisition for default full workflow

AI Agent integrations and skills MUST treat the full sales-report workflow as the default workflow. If the user provides two files without an explicit month, the Agent or Skill MUST try to infer `target_month` before running the CLI, and MUST ask the user when inference is not reliable.

#### Scenario: Infer month from filename or context
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** 文件名或聊天上下文包含可识别的月份信息（如 `202603`、`2026年3月`、`3月份`、`上个月`）
- **THEN** Agent/Skill 将月份归一化为 `YYYYMM`
- **AND** 使用该 `target_month` 调用完整工作流

#### Scenario: Ask user when month is missing
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** Agent/Skill 无法从文件名或上下文可靠识别月份
- **THEN** Agent/Skill MUST ask the user which month to process before invoking the CLI
- **AND** MUST NOT silently fall back to `--match-only` or basic matching

#### Scenario: Explicit reduced workflow request
- **WHEN** 用户明确表示只需要匹配手续费、不要销售报表或不要账期标注
- **THEN** Agent/Skill MAY run `--match-only`
- **AND** 该行为 MUST be treated as an explicit reduced workflow, not the default

### Requirement: CLI execution routes through workflow service

`cli.py` MUST route validated workflow execution through the workflow/service layer while preserving the existing CLI input contract.

#### Scenario: Full workflow route
- **WHEN** `cli.py` has valid `order_file`, `payment_file`, and `target_month` with no reduced mode flag
- **THEN** it SHALL call the full sales-report workflow service operation

#### Scenario: Match-only route
- **WHEN** `cli.py` has valid arguments and `--match-only`
- **THEN** it SHALL call the match-only workflow service operation

#### Scenario: Mark-only route
- **WHEN** `cli.py` has valid arguments and `--mark-only`
- **THEN** it SHALL call the mark-only workflow service operation

#### Scenario: Input behavior preserved
- **WHEN** `cli.py` routes execution through the service layer
- **THEN** positional argument parsing, target-month validation, interactive target-month prompting, and mode validation SHALL remain compatible with the existing CLI input requirements

### Requirement: 非交互式运行模式

excel_merge.py MUST 支持非交互式运行模式，允许 AI Agent 和自动化脚本在无 TTY 环境下调用，无需人工输入。

#### Scenario: 通过参数指定文件（非交互）
- **WHEN** 用户执行 `python excel_merge.py --order-file order.xlsx --payment-file payment.xlsx --non-interactive`
- **THEN** 程序直接使用指定的文件进行处理
- **AND** 不调用 `input()` 提示用户选择
- **AND** 处理完成后正常退出

#### Scenario: 自动检测非 TTY 环境
- **WHEN** 程序在无 TTY 的环境中运行（如 `echo "" | python excel_merge.py --order-file order.xlsx --payment-file payment.xlsx`）
- **THEN** 程序自动切换到非交互式模式
- **AND** 不调用 `input()` 提示用户选择

#### Scenario: 非交互模式下缺少必要参数
- **WHEN** 用户执行 `python excel_merge.py --non-interactive`（未指定文件）
- **THEN** 程序输出错误信息说明缺少 `--order-file` 和 `--payment-file` 参数
- **AND** 进程退出码为 2

#### Scenario: 交互模式保持不变
- **WHEN** 用户在有 TTY 的终端中执行 `python excel_merge.py`（不带 `--non-interactive`）
- **THEN** 程序行为与当前版本一致
- **AND** 列出 `ExcelForHandel/` 目录中的文件供用户选择

### Requirement: utils.py 日志系统统一

utils.py 中的所有用户可见输出 MUST 使用 Python `logging` 模块而非 `print()` 函数，以便调用方控制日志级别和输出目标。入口脚本（`cli.py`、`excel_merge.py`、`excel_merge_api.py`）MAY 使用 `print()` 输出其自身的用户提示、文本模式摘要或开发日志。

#### Scenario: logging 替代 print
- **WHEN** utils.py 中的任何函数被调用
- **THEN** 所有进度信息、调试信息、警告信息通过 `logging` 模块输出
- **AND** utils.py 不直接调用 `print()` 函数（注释和文档字符串中的除外）

#### Scenario: 日志级别映射
- **WHEN** `process_excel_files(verbose=True)` 被调用
- **THEN** 匹配过程的逐行详情使用 `logger.debug()` 级别
- **AND** 匹配摘要信息使用 `logger.info()` 级别
- **AND** 数据异常警告使用 `logger.warning()` 级别

#### Scenario: 日志级别映射（verbose=False）
- **WHEN** `process_excel_files(verbose=False)` 被调用
- **THEN** 仅输出 `logger.info()` 及以上级别的日志
- **AND** 不输出逐行匹配详情

### Requirement: excel_merge.py JSON 输出支持

excel_merge.py 在非交互式模式下 SHALL 支持 `--json` 标志，输出与 cli.py 相同格式的 JSON 结构化结果。

#### Scenario: 非交互 JSON 输出
- **WHEN** 用户执行 `python excel_merge.py --order-file order.xlsx --payment-file payment.xlsx --non-interactive --json`
- **THEN** stdout 输出有效 JSON，格式与 cli.py 的 `--json` 输出一致
- **AND** `ok` 为 `true`，`data` 包含处理结果统计

#### Scenario: 交互模式下忽略 JSON 标志
- **WHEN** 用户在有 TTY 的终端中执行 `python excel_merge.py --json`（未使用 `--non-interactive`）
- **THEN** 程序正常进入交互式文件选择流程
- **AND** 处理完成后以 JSON 格式输出结果

### Requirement: Interactive mode sales report trigger
The interactive mode (`excel_merge.py`) SHALL provide an option to trigger the sales report workflow after file selection.

#### Scenario: User opts to generate a sales report
- **WHEN** the user successfully selects an order and payment file in interactive mode
- **AND** the system prompts "Do you want to generate a sales report? (y/n)"
- **AND** the user enters 'y'
- **THEN** the system SHALL prompt the user to "Enter the report month (e.g., 202602): "
- **AND** the `process_sales_report_workflow` SHALL be called with the provided month.

#### Scenario: User declines to generate a sales report
- **WHEN** the user successfully selects an order and payment file in interactive mode
- **AND** the system prompts "Do you want to generate a sales report? (y/n)"
- **AND** the user enters 'n'
- **THEN** the standard file processing SHALL continue without triggering the sales report workflow.

#### Scenario: Invalid month format in interactive mode
- **WHEN** the user provides an invalid month format (e.g., "2026-02" or "abc")
- **THEN** the system SHALL display an error message and prompt again for a valid `YYYYMM` format.
