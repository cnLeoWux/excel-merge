# Change: 优化 CLI 接口以支持 AI Agent 调用

## Why

当前 CLI（cli.py）和交互式入口（excel_merge.py）的输出全部为人类可读的 print() 文本，混合在 stdout 上，没有结构化 JSON 模式、没有语义化退出码、没有静默模式。交互式入口使用阻塞式 `input()` 提示，在非交互环境下（CI/CD、AI Agent 如 OpenClaw）完全无法使用。这些问题使得 AI Agent 无法可靠地调用本工具、解析结果、判断成败并做出下一步决策。

Flask API（excel_merge_api.py）已经提供了结构化 JSON 响应，证明核心逻辑可以产出结构化数据。本提案将同样的能力扩展到 CLI 入口。

## What Changes

### CLI 结构化输出（cli.py）
- 新增 `--json` 标志：启用时所有结果以 JSON 格式输出到 stdout
- 新增 `--quiet` / `--verbose` 标志：控制日志详细程度（替代当前强制 `verbose=True`）
- **BREAKING**：引入语义化退出码（0=成功, 1=通用错误, 2=参数/用法错误, 3=文件未找到, 4=处理错误），替代当前所有错误路径返回退出码 0 的行为
- stdout/stderr 分离：数据（JSON 或最终结果路径）输出到 stdout，日志/进度/警告输出到 stderr

### 非交互式模式（excel_merge.py）
- 新增 `--non-interactive` 标志或自动检测 TTY 环境
- 新增 `--order-file` / `--payment-file` 参数，允许直接指定文件路径而非交互式选择
- 在非 TTY 环境下自动跳过 `input()` 提示

### 日志系统重构（utils.py）
- 将 `print()` 调用迁移到 Python `logging` 模块
- 日志输出到 stderr（不干扰 stdout 的结构化数据）
- 保留 `verbose: bool = False` 参数模式，但日志级别由调用方（CLI 标志）控制

### 错误处理增强
- 结构化错误响应（JSON 模式下返回包含 `code`、`message` 字段的错误对象）
- 用 `sys.exit(code)` 替代当前的 `print() + return` 错误处理模式

## 非目标（Non-goals）

- **不重写匹配算法**：匹配逻辑（精确匹配、P-number、连字符）保持不变
- **不修改 Flask API**：API 已有 JSON 支持，不在本次改动范围内
- **不引入新依赖**：使用 Python 标准库（logging、json、sys）实现所有功能
- **不改变默认人类可读行为**：不带 `--json` 标志时，CLI 行为与当前一致（向后兼容）
- **不实现 CLI introspection（schema 命令）**：暂不引入自描述能力，留待后续迭代
- **不实现 dry-run 模式**：预览功能不在本次范围内
- **不实现请求去重（idempotency）**：留待后续需求驱动

## Impact

### 受影响的规格
- 新增能力：`cli-output`（CLI 结构化输出）
- 新增能力：`cli-input`（CLI 非交互式输入）

### 受影响的代码

| 文件 | 受影响函数/区域 | 改动类型 |
|------|----------------|----------|
| `cli.py` | `main_cli()` — argparse 参数定义（L21-49）| 新增 `--json`、`--quiet`、`--verbose` 参数 |
| `cli.py` | `main_cli()` — 文件存在性检查（L54-60）| 替换 `print()+return` 为 `sys.exit(3)` |
| `cli.py` | `main_cli()` — 处理结果输出（L85-113）| 新增 JSON 输出路径 |
| `cli.py` | `main_cli()` — 异常处理（L115-119）| 新增 JSON 错误输出 + `sys.exit(4)` |
| `excel_merge.py` | `main()` — 文件选择逻辑（L40, L52）| 新增非交互式路径，跳过 `input()` |
| `utils.py` | `process_excel_files()` 内所有 `print()` 调用 | 迁移到 `logging` 模块 |
| `utils.py` | `add_sales_report_period()` 内所有 `print()` 调用 | 迁移到 `logging` 模块 |
| `utils.py` | `filter_unmarked_and_generate_report()` 内所有 `print()` 调用 | 迁移到 `logging` 模块 |
| `utils.py` | `read_file_with_appropriate_method()` 内 `print()` 调用 | 迁移到 `logging` 模块 |
| `utils.py` | `process_sales_report_workflow()` 内 `print()` 调用 | 迁移到 `logging` 模块 |
| `setup.py` | `entry_points` | 无改动（console_scripts 入口不变）|
