## 为什么

workflow/service 层已经开始集中编排，但一些边界职责仍然分散在各个 adapter 中：`target_month`/month 校验、file-not-found 处理、processing-error 映射，以及部分文档仍在描述旧的 `--month` CLI 形态。在拆分 `utils.py` 之前先加固这条边界，可以降低后续模块重构时把不一致的错误与文档行为继续带入的风险。

## 变更内容

- 强化 service 级输入校验：让无效或缺失的 `target_month`/API `month` 值在核心 workflow 执行前就转换为规范化的 `WorkflowError(code="usage_error")`。
- 为 `WorkflowError` 路径补充聚焦测试：缺失文件、无效月份、写入失败，以及 API service metadata 行为。
- 将更多 workflow 失败分类下沉到 service 层，让 CLI/API adapter 主要负责格式化已规范化的错误。
- 通过把当前 CLI `--month` 示例替换为位置参数 `target_month` 示例来修正文档漂移，并明确 `--match-only` 是显式缩减工作流。
- 保持公开的 CLI/API envelope 和持久化行为不变；不改变匹配或销售报表算法。

## Capabilities

### 新能力

- 无。

### 修改的能力

- `workflow-service`：收紧 service 校验/错误规范化、API 报表统计与 service metadata 预期。
- `cli-output`：明确 CLI 格式化消费的是已规范化的 service 错误，同时保留现有退出码与 JSON envelope。
- `http-api`：明确 API month 校验以及 workflow 失败时的 service-error-to-HTTP 映射。
- `agent-documentation`：将 AGENTS.md 和面向用户的示例对齐到位置参数 `target_month` 与“默认完整工作流”的表述。
- `automated-testing`：要求 service 错误规范化测试与修正后 CLI 示例的文档契约测试。

## 影响

- 受影响代码：`workflow_service.py`、`cli.py`、`excel_merge_api.py`，以及必要时 `excel_merge.py` 中的小型 adapter 清理。
- 受影响测试：`tests/unit/test_workflow_service.py`、CLI/API 集成测试，以及文档/Skill 断言测试。
- 受影响文档：`AGENTS.md`、`documents/USAGE_EXAMPLES.md`，以及可能的 `.opencode/skills/excel-merge-cli/SKILL.md` 文案。
- 本次变更不引入新的运行时依赖，也不拆分 `utils.py` 模块。
