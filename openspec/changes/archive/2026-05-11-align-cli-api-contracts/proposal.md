## 为什么

CLI、Agent Skill、交互入口与 HTTP API 目前暴露了多套相近但不一致的契约：CLI 使用 `ok/data/error` JSON 信封，API 使用 `success/download_url/files` 形状；CLI 默认完整流程依赖 `target_month`，而文档与自动化调用方式容易被误解为可默认基础匹配。现在先对齐这些入口契约，避免后续抽 workflow/service 层或拆分 `utils.py` 时把不一致固化到更多模块。

## 变更内容

- 统一默认自动化意图：当 `target_month` 可用或可被获取时，两份已上传/已提供文件 SHOULD 运行完整销售报表工作流；`--match-only` 是显式的缩减工作流。
- 明确 Agents 和 Skills 如何获取 `target_month`：在可靠时从文件名/对话中推断，否则在调用 CLI 前询问用户。
- 围绕当前完整工作流的 `target_month --json --quiet` 调用对齐 CLI JSON 与文档，并在完整工作流统计中包含 `marked_rows`。
- 明确并文档化 CLI JSON 与 HTTP API JSON 的关系，而不是让 `/merge/json` 与 CLI 保持相互矛盾的契约。
- 让 HTTP API 请求/响应行为与所选契约对齐，包括文件校验、错误形状、销售报表行为与可下载产物。
- 明确 CLI 与 API 的文件输出语义：CLI 直接就地写回订单文件，不生成 `report_*.xlsx`；如果仍采用该 API 契约，API MAY 将可下载结果/报表文件持久化到 `results/`。
- 更新面向 Agent 的文档和测试，使其验证同一套 CLI/API 契约，而不是保留旧的冲突假设。

## Capabilities

### 新能力

- 无。

### 修改的能力

- `cli-input`：定义标准完整工作流调用、`target_month` 获取预期与显式缩减模式行为。
- `cli-output`：将 JSON 统计、stdout/stderr 预期、取消行为与错误码对齐到所选 CLI 契约。
- `http-api`：将 `/merge`、`/merge/json`、错误响应、文件校验、MIME/download 行为与销售报表产物对齐到所选 API 契约。
- `agent-documentation`：更新面向 Agent/Skill 的使用规则，使自动化在默认完整工作流中获取月份，并且不会静默回退到仅匹配模式。
- `automated-testing`：更新预期的 CLI/API 测试行为，以验证对齐后的契约。
- `sales-report`：明确 CLI、交互模式与 API 如何触发完整工作流，以及 API 专属报表持久化如何与核心工作流关联。

## 影响

- 受影响的入口：`cli.py`、`excel_merge.py`、`excel_merge_api.py`。
- 受影响的 Agent 面：`AGENTS.md`、`.opencode/skills/excel-merge-cli/SKILL.md` 以及面向用户的使用文档。
- 受影响的测试：CLI 子进程/主流程测试、Flask API 集成测试，以及任何断言 JSON 形状或报表文件行为的测试。
- 受影响的规格：`cli-input`、`cli-output`、`http-api`、`agent-documentation`、`automated-testing` 与 `sales-report`。
- 预期不会新增运行时依赖；变更应先聚焦于契约对齐与兼容性决策，再进行更广泛的 workflow/service 重构。
