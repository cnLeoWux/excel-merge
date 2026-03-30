## 1. 日志基础设施（utils.py）

受影响函数：`read_file_with_appropriate_method()`、`process_excel_files()`、`add_sales_report_period()`、`filter_unmarked_and_generate_report()`、`process_sales_report_workflow()`

- [x] 1.1 在 utils.py 顶部配置 `logger = logging.getLogger(__name__)`（utils.py 已 import logging）
- [x] 1.2 将 `read_file_with_appropriate_method()` 中的 `print()` 替换为 `logger.warning()` / `logger.debug()`（约 1 处，L87）
- [x] 1.3 将 `process_excel_files()` 中的所有 `print()` 替换为 `logger.info()` / `logger.debug()`（约 30+ 处，L205-L512）
- [x] 1.4 将 `add_sales_report_period()` 中的 `print()` 替换为 `logger.info()` / `logger.debug()`（约 15 处，L609-L682）
- [x] 1.5 将 `filter_unmarked_and_generate_report()` 中的 `print()` 替换为 `logger.info()` / `logger.warning()`（约 15 处，L780-L884）
- [x] 1.6 将 `process_sales_report_workflow()` 中的 `print()` 替换为 `logger.info()`（约 5 处，L913-L918）
- [x] 1.7 验证：运行 `grep -n "print(" utils.py` 确认无遗漏的 print 调用（允许保留注释中的 print）

## 2. CLI 结构化输出（cli.py）

受影响函数：`main_cli()`

- [x] 2.1 新增 argparse 参数：`--json`（action="store_true"）、`--verbose` / `-v`（action="count", default=0）、`--quiet`（action="store_true"）
- [x] 2.2 在 `main_cli()` 开头配置 logging：根据 `--verbose` / `--quiet` 设置日志级别，handler 输出到 stderr
- [x] 2.3 实现 `output_result(data, error, json_mode)` 辅助函数：JSON 模式输出信封到 stdout，文本模式保持原有 print 行为
- [x] 2.4 替换文件存在性检查（L54-60）：改用 `output_result()` + `sys.exit(3)`
- [x] 2.5 替换正常处理结果输出（L85-113）：改用 `output_result()`，包含 `output_file` 和 `statistics`
- [x] 2.6 替换异常处理（L115-119）：改用 `output_result()` + `sys.exit(4)`
- [x] 2.7 将 `verbose=True` 硬编码（L73, L100）改为根据 `--verbose` / `--quiet` 标志动态设置
- [x] 2.8 验证：手动运行以下命令并检查输出
  - `python cli.py nonexistent.xlsx payment.xlsx --json` → JSON 错误 + 退出码 3
  - `python cli.py order.xlsx payment.xlsx --json` → JSON 成功 + 退出码 0
  - `python cli.py order.xlsx payment.xlsx --quiet` → 无进度日志
  - `python cli.py order.xlsx payment.xlsx` → 原有行为不变
- [x] 3.1 在 cli.py 顶部定义退出码常量：`EXIT_SUCCESS = 0`、`EXIT_GENERAL_ERROR = 1`、`EXIT_USAGE_ERROR = 2`、`EXIT_FILE_NOT_FOUND = 3`、`EXIT_PROCESSING_ERROR = 4`
- [x] 3.2 在所有错误路径上使用 `sys.exit(code)` 替代 `return`
- [x] 3.3 验证：运行 `python cli.py nonexistent.xlsx payment.xlsx; echo $?` → 输出 3

## 4. 非交互式模式（excel_merge.py）

受影响函数：`main()`

- [x] 4.1 新增 argparse 解析器，定义参数：`--order-file`、`--payment-file`、`--non-interactive`、`--json`、`--output`
- [x] 4.2 实现 TTY 检测逻辑：`args.non_interactive or not sys.stdin.isatty()`
- [x] 4.3 在非交互模式下跳过 `input()` 调用，直接使用 `--order-file` / `--payment-file` 指定的文件
- [x] 4.4 如果非交互模式下未提供必要的文件参数，输出错误并 `sys.exit(2)`
- [x] 4.5 验证：
  - `python excel_merge.py --non-interactive --order-file order.xlsx --payment-file payment.xlsx` → 正常处理
  - `echo "" | python excel_merge.py` → 错误提示（非交互但未提供文件）

## 5. 文档更新

- [x] 5.1 更新 README.md：新增 Agent 调用示例（`--json`、`--quiet`、退出码表）
- [x] 5.2 更新 USAGE.md：新增非交互式模式和 JSON 输出说明
- [x] 5.3 更新 documents/USAGE_EXAMPLES.md：新增 Agent 场景的用法示例
- [x] 5.4 验证：检查文档中的示例命令可正常执行
