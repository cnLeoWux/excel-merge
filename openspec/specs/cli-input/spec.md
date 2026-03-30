## Purpose

CLI 输入能力 - 定义命令行工具的非交互式运行模式，支持 AI Agent 和自动化脚本在无 TTY 环境下调用。

## Requirements

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

### Requirement: 日志系统统一

utils.py 中的所有用户可见输出 MUST 使用 Python `logging` 模块而非 `print()` 函数，以便调用方控制日志级别和输出目标。

#### Scenario: logging 替代 print
- **WHEN** utils.py 中的任何函数被调用
- **THEN** 所有进度信息、调试信息、警告信息通过 `logging` 模块输出
- **AND** 不直接调用 `print()` 函数（注释和文档字符串中的除外）

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
