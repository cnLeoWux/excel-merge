## Context

本项目有三个入口（cli.py、excel_merge.py、excel_merge_api.py），核心逻辑集中在 utils.py（~930 行）。Flask API 已提供结构化 JSON 响应，但 CLI 和交互式入口完全依赖 `print()` 输出到 stdout，没有机器可读格式。

AI Agent（如 OpenClaw、Cursor、Claude Code）调用 CLI 工具时需要：
1. 结构化输出（JSON）以便解析结果
2. 语义化退出码以判断成败和错误类型
3. stdout/stderr 分离以便只解析 stdout 数据
4. 非交互式运行（无 `input()` 阻塞）
5. 可控的日志级别（静默运行避免浪费 token）

### 利益相关方
- AI Agent 开发者（OpenClaw 等）：需要可靠的 CLI 调用接口
- 现有人类用户：需要保持向后兼容的默认行为
- 项目维护者：需要统一的日志基础设施

## Goals / Non-Goals

### Goals
1. CLI 支持 `--json` 标志，输出结构化 JSON 到 stdout
2. CLI 支持 `--quiet` 和 `--verbose` 标志控制日志级别
3. 引入语义化退出码（0/1/2/3/4）
4. stdout 仅输出数据，stderr 承载日志/进度/警告
5. excel_merge.py 支持非交互式模式
6. utils.py 的 `print()` 迁移到 `logging` 模块

### Non-Goals
- 不重写匹配算法或改变业务逻辑
- 不修改 Flask API（已有 JSON 支持）
- 不引入 click、rich 等新依赖
- 不实现 CLI introspection/schema 命令
- 不实现 dry-run 或请求去重功能

## Decisions

### 决策 1：JSON 输出格式采用统一信封（Envelope）

**选择**：所有 JSON 输出使用统一的顶层结构

```json
{
  "ok": true,
  "data": {
    "output_file": "result.xlsx",
    "statistics": {
      "total_rows": 100,
      "matched_rows": 85,
      "match_rate": "85.00%"
    }
  },
  "error": null
}
```

错误时：
```json
{
  "ok": false,
  "data": null,
  "error": {
    "code": "file_not_found",
    "message": "File 'order.xlsx' does not exist"
  }
}
```

**替代方案**：
- 复用 Flask API 的响应格式（含 `success`、`session_id`、`download_url`）→ 不适合 CLI 场景，CLI 无需 session_id 或 download_url
- 直接输出裸 JSON 数据（无信封）→ Agent 无法区分成功和失败

**理由**：统一信封让 Agent 可以先检查 `ok` 字段判断成败，再处理 `data` 或 `error`。与业界最佳实践（gh CLI、kubectl）一致。

### 决策 2：退出码定义

| 退出码 | 含义 | 场景 |
|--------|------|------|
| 0 | 成功 | 处理完成，结果已写入 |
| 1 | 通用错误 | 未预期的异常 |
| 2 | 参数/用法错误 | argparse 已处理；保留给无效参数 |
| 3 | 文件未找到 | 输入文件不存在 |
| 4 | 处理错误 | 匹配或写入过程中的业务错误 |

**替代方案**：
- 只用 0 和 1 → Agent 无法区分错误类型，无法决定是否重试
- 使用 HTTP 风格的状态码（200/404/500）→ 不符合 Unix CLI 惯例（退出码范围 0-255）

**理由**：5 个退出码覆盖所有常见场景，Agent 可根据退出码决定重试策略。argparse 的默认退出码 2 已符合此方案。

### 决策 3：日志迁移策略

**选择**：渐进式迁移 — 在 utils.py 中用 `logging.getLogger(__name__)` 替换 `print()`，CLI 入口配置 logging handler

```python
# utils.py 中
import logging
logger = logging.getLogger(__name__)

# 原来：print("Starting matching process...")
# 改为：logger.info("Starting matching process...")

# cli.py 中
import logging
logging.basicConfig(
    level=logging.WARNING,  # 默认只输出警告
    format="%(message)s",
    stream=sys.stderr       # 日志到 stderr
)
if args.verbose:
    logging.getLogger().setLevel(logging.DEBUG)
elif not args.quiet:
    logging.getLogger().setLevel(logging.INFO)
```

**替代方案**：
- 引入 structlog → 增加新依赖，违反 Non-goals
- 保留 print() 但重定向到 stderr → 不够灵活，无法按级别过滤

**理由**：Python 标准 `logging` 模块零依赖，与 `verbose: bool` 参数模式兼容。utils.py 已 import logging（未使用），迁移成本低。

### 决策 4：非交互式模式实现

**选择**：excel_merge.py 新增 argparse 参数 + TTY 自动检测

```python
# 新增参数
parser.add_argument("--order-file", type=str, help="直接指定订单文件路径")
parser.add_argument("--payment-file", type=str, help="直接指定支付文件路径")
parser.add_argument("--non-interactive", action="store_true")

# 自动检测
if args.non_interactive or not sys.stdin.isatty():
    # 使用参数指定的文件，不调用 input()
```

**替代方案**：
- 只检测 TTY，不加参数 → 无法在有 TTY 的环境下强制非交互
- 使用环境变量 `NON_INTERACTIVE=1` → 不够显式，容易遗漏

**理由**：参数 + TTY 检测双保险，确保 AI Agent 在任何环境下都能非交互运行。

## Risks / Trade-offs

| 风险 | 缓解措施 |
|------|---------|
| utils.py 的 print→logging 迁移可能遗漏某些 print 调用 | 用 `grep -n "print(" utils.py` 做全量检查，迁移后运行全流程验证 |
| 退出码变更可能影响已有的 shell 脚本调用 | 仅在 **错误路径** 上变更退出码（原来也是 0，脚本不太可能依赖错误时的退出码为 0） |
| `--json` 模式下 utils.py 仍可能有未迁移的 print 泄漏到 stdout | 在 `--json` 模式下增加 stdout 拦截层：将 sys.stdout 临时替换为 buffer，最后只输出 JSON |
| excel_merge.py 新增 argparse 后与 cli.py 功能重叠 | 明确定位：excel_merge.py 面向"从 ExcelForHandel/ 目录选文件"的场景，cli.py 面向"指定任意文件路径"的通用场景 |

## Migration Plan

### 阶段 1：非破坏性新增（向后兼容）
1. utils.py：新增 logging 调用，保留 print（双写过渡期）
2. cli.py：新增 `--json`、`--quiet`、`--verbose` 参数（默认行为不变）
3. excel_merge.py：新增 `--non-interactive`、`--order-file`、`--payment-file` 参数

### 阶段 2：完成迁移
4. utils.py：移除所有 print() 调用，全部使用 logging
5. cli.py：在错误路径上使用 `sys.exit(code)`
6. 更新 README.md 和 USAGE.md 中的 CLI 文档

### 回滚方案
- 所有改动通过 git 管理，可随时 revert 单个 commit
- 新增的 `--json` 等参数为可选，不影响无参数调用

## Open Questions

1. **JSON 输出中是否需要包含 `schemaVersion` 字段？** — 可在 v2 时添加，v1 暂不需要
2. **sales report 工作流的 JSON 输出应包含哪些额外字段？** — 建议增加 `report_file` 和 `report_rows` 字段
3. **是否需要支持 `--json` 与 `--month` 的组合？** — 建议支持，返回完整工作流结果
