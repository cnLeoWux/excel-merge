# Tasks

## 1. 行为锁定测试

> 注：本阶段先“证明当前行为”，不要借测试整理顺手改变业务语义。尤其是 fallback 行顺序、`process_excel_files()` 刷新 `销售报表账期`、CSV fallback 顺序，都应先按现状写成 golden tests。

- [ ] 1.1 为 `process_excel_files()` 增加或确认 golden tests：20 字符 exact 优先、P-number fallback、hyphen fallback、fallback payment 行顺序。
- [ ] 1.2 为正单、退单、零金额三种订单金额路径增加或确认测试。
- [ ] 1.3 增加或确认测试：`process_excel_files()` 返回结果仍包含并刷新 `销售报表账期`。
- [ ] 1.4 为 `read_file_with_appropriate_method()` 增加或确认 CSV 编码 fallback、分隔符 fallback、注释行跳过、订单号字符串保护测试。
- [ ] 1.5 为 Excel `.xlsx` / `.xls` 读取路径增加或确认测试。
- [ ] 1.6 运行 `python -m pytest tests/unit -v`。

## 2. 在 `utils.py` 内提取小函数

受影响函数：`process_excel_files()`、`extract_p_number()`、`match_orders_by_p_number()`。

> 注：先在原文件内提取 helper，是为了降低一次性移动代码造成的风险。每完成一个 helper 提取，都应能用第 1 阶段测试证明可观察结果没有变化。

- [ ] 2.1 提取业务订单号列识别 helper。
- [ ] 2.2 提取订单金额方向分类 helper。
- [ ] 2.3 提取业务类型兼容性 helper。
- [ ] 2.4 提取支付手续费金额选择 helper。
- [ ] 2.5 提取 exact 匹配和 fallback 匹配 helper。
- [ ] 2.6 运行核心匹配相关测试，确认行为不变。

## 3. 拆分文件 I/O 模块

受影响函数：`read_file_with_appropriate_method()`、`write_result_file()`、`find_file_path()`。

> 注：拆分时优先移动纯 I/O 与归一化逻辑；不要把支付匹配、销售报表标注或 service 统计顺手放入 `file_io.py`。

- [ ] 3.1 新增 `file_io.py`，迁移 CSV 读取 fallback 和注释行处理。
- [ ] 3.2 迁移 Excel 引擎检测逻辑。
- [ ] 3.3 迁移订单号/流水号列字符串保护和清理逻辑。
- [ ] 3.4 迁移写回和文件查找逻辑。
- [ ] 3.5 让 `utils.py` 继续导出旧 I/O 函数名。
- [ ] 3.6 运行文件 I/O 相关测试。

## 4. 拆分匹配模块

受影响函数：`process_excel_files()`、`extract_p_number()`、`match_orders_by_p_number()`。

> 注：`matching.py` 只关心订单行、支付行和匹配结果。CLI 参数、HTTP 请求、下载文件名、JSON envelope 都不应进入该模块。

- [ ] 4.1 新增 `matching.py`，迁移 P-number 和匹配 helper。
- [ ] 4.2 保持 `process_excel_files()` 外部签名和返回行为不变。
- [ ] 4.3 确认 fallback 匹配仍按 payment 文件行顺序扫描。
- [ ] 4.4 让 `utils.py` 继续导出旧匹配函数名。
- [ ] 4.5 运行核心匹配和 workflow service 测试。

## 5. 拆分销售报表模块

受影响函数：`add_sales_report_period()`、`parse_date()`、`get_year_month()`、`filter_unmarked_and_generate_report()`、`process_sales_report_workflow()`。

> 注：保持“完整工作流不写独立报表文件”的当前契约。若实现中发现注释和实际日期窗口不一致，只记录风险，不在本 change 中改语义。

- [ ] 5.1 新增 `sales_report.py`，迁移销售报表账期标注逻辑。
- [ ] 5.2 迁移日期解析和年月提取逻辑。
- [ ] 5.3 迁移未标注行筛选和完整销售报表工作流。
- [ ] 5.4 保持完整工作流不写出独立报表文件。
- [ ] 5.5 让 `utils.py` 继续导出旧销售报表函数名。
- [ ] 5.6 运行销售报表相关测试。

## 6. 整理 workflow/service 与 adapter 边界

受影响函数/文件：`workflow_service.py`、`cli.py`、`excel_merge.py`、`excel_merge_api.py`。

> 注：service 负责应用级编排和统计复用；adapter 负责用户可见的 transport contract。不要为了复用而把 CLI JSON envelope 套到 API 上，也不要让 API response shape 影响 CLI。

- [ ] 6.1 确认 `workflow_service.py` 只负责编排、统计、持久化协调和错误归一化。
- [ ] 6.2 确认 CLI 输出使用 service result，不重复计算共享统计。
- [ ] 6.3 确认 API 路由保留原响应 shape，不复用 CLI JSON envelope。
- [ ] 6.4 去除明显重复的 adapter 辅助逻辑，但不改变外部输出字段。
- [ ] 6.5 运行 CLI 和 API 集成测试。

## 7. 全量验证与文档

- [ ] 7.1 运行 `python -m pytest`。
- [ ] 7.2 运行 `openspec validate refactor-core-workflow-boundaries --strict`。
- [ ] 7.3 检查 `documents/` 和 README 中是否有与新模块边界冲突的描述。
- [ ] 7.4 如实现后模块边界稳定，更新架构文档中的代码结构说明。
