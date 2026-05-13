# Proposal: 重构核心工作流边界

## 背景

当前项目的核心业务逻辑集中在 `utils.py` 中，文件读取、支付手续费匹配、销售报表账期标注、月度报表筛选、文件写回等职责彼此交织。`cli.py`、`excel_merge.py`、`excel_merge_api.py` 和 `workflow_service.py` 已经开始共享工作流层，但核心模块仍然难以测试、难以安全调整。

这次变更的重点不是改变业务行为，而是为后续维护建立清晰边界：先用测试钉住现有契约，再拆分职责和小函数，最后让 CLI/API/服务层更稳定地复用结果对象。

## 问题

- `utils.py` 近千行，混合了文件 I/O、匹配算法、销售报表、日期解析和写回逻辑。
- `process_excel_files()` 同时负责匹配、统计所需列生成，并隐式刷新 `销售报表账期`，副作用重要但不直观。
- `read_file_with_appropriate_method()` 内部包含多层编码、分隔符和 Excel 引擎 fallback，重复清理逻辑较多，错误边界不清。
- CLI/API/service 对统计、错误、输出数据有重复拼装，长期容易出现契约漂移。
- 性能优化和包结构迁移都很诱人，但如果没有 golden tests，很容易改变匹配优先级或文件读写行为。

## 目标

1. 在不改变外部行为的前提下，将核心逻辑拆成更清晰的职责边界。
2. 保留 `utils.py` 作为兼容 facade，使现有入口和测试可以渐进迁移。
3. 将 `process_excel_files()` 拆成可测试的小函数，同时保留现有匹配顺序、副作用和列语义。
4. 将 CSV/Excel 文件读取和写入逻辑拆成独立 I/O 组件，同时保留编码、分隔符和引擎 fallback 顺序。
5. 明确 workflow/service 层、CLI/API adapter 和核心业务模块之间的职责分工。
6. 增加行为锁定测试，尤其覆盖 exact / P-number / hyphen / fallback 行顺序、正单/退单/零金额、CSV/Excel 读取。

## 非目标

- 不改变匹配算法的业务结果或优先级。
- 不将 P-number 提取改为大小写不敏感。
- 不移除 `process_excel_files()` 当前刷新 `销售报表账期` 的副作用。
- 不改变 CLI JSON envelope、退出码、stdout/stderr 约定。
- 不统一 CLI 与 HTTP API 的 JSON 响应 shape。
- 不迁移到 `src/` 包结构。
- 不重写 Flask API 为 app factory。
- 不进行大规模性能优化或索引化匹配。

## 影响分析

### `utils.py`

- `extract_p_number()`、`match_orders_by_p_number()`、`process_excel_files()`、`read_file_with_appropriate_method()`、`write_result_file()`、`find_file_path()`、`add_sales_report_period()`、`filter_unmarked_and_generate_report()`、`process_sales_report_workflow()` 的外部调用方式必须保持兼容。
- 可将实现迁移到新模块，但 `utils.py` 必须继续暴露旧函数名。
- `process_excel_files()` 的返回 DataFrame 必须仍包含 `支付手续费` 与刷新后的 `销售报表账期`。

### 新核心模块

- `file_io.py`：承载 CSV/Excel 读取、编码/分隔符 fallback、订单号字符串保护、结果写回、文件查找。
- `matching.py`：承载 P-number、业务类型校验、金额赋值、精确匹配、fallback 匹配等纯匹配逻辑。
- `sales_report.py`：承载账期标注、日期解析、未标注行筛选与完整销售报表工作流。

### `workflow_service.py`

- 保持应用级编排职责：校验、调用核心工作流、持久化协调、统计生成、错误归一化。
- 不应吸收核心匹配算法或文件读取 fallback 细节。

### `cli.py` / `excel_merge.py` / `excel_merge_api.py`

- 保持 adapter 职责：参数解析、交互、HTTP 请求处理、输出格式、退出码/HTTP 状态映射。
- 不应重复计算共享统计或复制核心业务分支。

### 测试

- 增加或调整 `tests/` 下的 pytest 测试，锁定行为后再拆代码。
- 根目录的手动验证脚本不作为本变更的主要验证基础。

## 风险

- 拆分过程中可能无意改变 fallback 匹配的行顺序优先级。
- 文件读取 fallback 顺序如果变化，会影响历史 CSV/Excel 文件兼容性。
- `process_excel_files()` 的销售报表账期副作用容易在“纯化”过程中被误删。
- CLI/API 输出字段若顺手整理，可能破坏已有自动化调用方。

## 迁移策略

采用“先测试、再提取、后迁移调用”的渐进式方式：

1. 先补 golden tests 和服务/adapter 契约测试。
2. 在 `utils.py` 内提取小函数，确保行为不变。
3. 新增 `file_io.py`、`matching.py`、`sales_report.py`，将实现迁入新模块。
4. `utils.py` 保持兼容导入/转发。
5. 再逐步让 service 和 adapter 直接依赖更清晰的边界。

### 兼容迁移说明

- 迁移期间不要要求调用方一次性改 import；旧代码继续从 `utils.py` 导入，新增测试可以直接覆盖新模块。
- `utils.py` 的 facade 只做导入/转发，不重新实现业务分支，避免“旧路径”和“新路径”出现两套行为。
- 如果某个 helper 在拆分后需要被多个模块复用，优先放在职责最贴近的 core module 中，而不是放回 `utils.py`。
- 本 change 的验收重点是“可观察行为不变”；模块数量、helper 名称和内部组织可以按实现中发现的约束微调。

## 成功标准

- `python -m pytest` 通过。
- `openspec validate refactor-core-workflow-boundaries --strict` 通过。
- 现有 CLI/API/OpenSpec 契约不发生破坏性变化。
- 核心职责边界在代码结构上可见，`utils.py` 不再承载全部实现细节。
