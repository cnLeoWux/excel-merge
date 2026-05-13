## 为什么

CLI、交互入口和 HTTP API 目前各自编排文件读取、核心处理、统计、写回、错误输出和下载响应，导致相同行为在多个入口中重复且容易漂移。引入统一 workflow/service 层可以先稳定入口与业务编排边界，为后续拆分 `utils.py`、纯化匹配引擎和改进持久化策略降低风险。

## 变更内容

- 添加一个 workflow/service 层，对外提供稳定的应用级操作，用于匹配与销售报表工作流。
- 将共享的编排职责移出入口：统计构建、规范化结果对象、规范化错误对象，以及写回协调。
- 重构 `cli.py`、`excel_merge.py` 和 `excel_merge_api.py`，改为调用 workflow/service 层，而不是重复编排逻辑。
- 保持已对齐的 CLI/API 规格中的当前行为和公开契约，包括默认完整工作流语义、CLI 就地写回，以及 API 可下载的结果/报表文件。
- 保持 `utils.py` 业务函数可用且兼容；本次变更是对它们进行包装和协调，而不是拆分或重写匹配逻辑。
- 添加/调整测试，使每个入口验证 adapter 行为，同时共享 workflow 行为通过 service 层测试。

## Capabilities

### 新能力

- `workflow-service`：定义应用 service 层，协调核心文件处理、销售报表工作流执行、结果统计、持久化决策，以及面向入口的规范化成功/错误结果。

### 修改的能力

- `cli-input`：CLI 行为保持契约兼容，但执行路径改为经过 workflow/service 层。
- `cli-output`：CLI JSON/text 输出保持契约兼容，同时使用 workflow/service 结果对象来处理统计与错误映射。
- `http-api`：API 行为保持契约兼容，同时使用 workflow/service 操作执行匹配与销售报表处理。
- `sales-report`：销售报表工作流在语义上保持不变，但由入口通过 workflow/service 层调用。
- `automated-testing`：测试必须覆盖新的 workflow/service 层，并确保入口仍然是轻量 adapter。

## 影响

- 受影响代码：新的 workflow/service 模块、`cli.py`、`excel_merge.py`、`excel_merge_api.py`，以及 `utils.py` 导入中的少量兼容性辅助代码。
- 受影响测试：新增 workflow/service 行为的单元测试，以及 CLI 和 API 集成测试的更新。
- 受影响文档/规格：新的 `workflow-service` 能力以及相关入口能力的 delta specs。
- 预计不会引入新的运行时依赖。
- 匹配算法改动、`utils.py` 模块拆分和 API envelope 版本化均不在本次变更范围内。
