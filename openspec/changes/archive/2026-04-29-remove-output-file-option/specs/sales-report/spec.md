## MODIFIED Requirements

### Requirement: 月度报表筛选

`filter_unmarked_and_generate_report()` MUST 仅保留未标注（`销售报表账期` 为空）且 `出行日期` 落在目标月份前后 1 年窗口内的行，作为内存中的中间 DataFrame 返回，供调用方进一步处理。该函数 MUST NOT 将筛选结果写入任何文件，且 MUST NOT 接受 `output_dir` 参数。

#### Scenario: 仅保留未标注行
- **WHEN** 生成月度报表 DataFrame
- **THEN** 所有 `销售报表账期` 为 `"全退"` 或 `"已取消"` 的行被排除
- **AND** 仅未标注（NaN 或空）的行进入返回的 DataFrame

#### Scenario: 出行日期窗口
- **WHEN** 目标月份为 `202602`（2026 年 2 月）
- **THEN** 返回的 DataFrame 仅保留 `出行日期` ∈ [2025-02-01, 2027-02-28] 的行（目标月份前后 1 年）
- **AND** `出行日期` 缺失或无法解析的行被排除

#### Scenario: 不写出报表文件
- **WHEN** 任意调用 `filter_unmarked_and_generate_report(...)` 完成
- **THEN** 文件系统中不会新增任何 `report_*.xlsx` 或其它结果文件
- **AND** 函数仅返回 `(updated_df, filtered_report_df)` 元组
- **AND** 函数签名不接受 `output_dir` 形参

#### Scenario: 空筛选结果
- **WHEN** 筛选后无任何符合条件的行
- **THEN** 返回的 `filtered_report_df` 为空 DataFrame
- **AND** 不产生任何文件副作用

### Requirement: 端到端工作流编排

`process_sales_report_workflow()` MUST 编排完整两阶段流程：匹配支付手续费 → 标注账期 → 计算未标注的报表 DataFrame。工作流函数自身 MUST NOT 写入任何文件；它返回 `(updated_df, filtered_report_df)` 供调用方决定如何持久化。该工作流 MUST 可由 CLI、交互模式与 HTTP API 触发；CLI 与交互模式 MUST 将 `updated_df` 就地写回原始订单文件、不产生其它文件，HTTP API 的文件写入语义由 `http-api` capability 单独定义，本 capability 不约束。

#### Scenario: 完整工作流执行顺序（CLI）
- **WHEN** CLI 收到 `--month 202602` 参数
- **THEN** 依次执行：
  1. `process_excel_files()` 完成支付手续费匹配
  2. `add_sales_report_period()` 标注 `销售报表账期` 列
  3. `filter_unmarked_and_generate_report()` 计算筛选 DataFrame（不落盘）
  4. `write_result_file()` 将更新后的订单 DataFrame 就地写回原始订单文件
- **AND** 任一步骤失败终止后续步骤并向调用方传播异常
- **AND** 不在任何目录生成 `report_YYYYMM.xlsx` 等独立报表文件

#### Scenario: 完整工作流执行顺序（Interactive）
- **WHEN** 用户在交互模式下选择为某个月份生成销售报表
- **THEN** 执行的函数序列与 CLI 工作流相同
- **AND** 同样不产生独立的报表文件

#### Scenario: 完整工作流执行顺序（API）
- **WHEN** HTTP API 在 `/merge` 或 `/merge/json` 路由中以 `month` 参数触发该工作流
- **THEN** 同样调用 `process_excel_files()` → `add_sales_report_period()` → `filter_unmarked_and_generate_report()`
- **AND** 工作流函数自身不写出任何文件
- **AND** 后续的文件写入由 API 层处理（详见 `http-api` capability）

#### Scenario: 工作流函数签名
- **WHEN** 调用 `process_sales_report_workflow(order_file, payment_file, target_month, verbose=...)`
- **THEN** 函数不接受 `output_dir` 形参
- **AND** 函数返回 `(updated_df, filtered_report_df)` 元组供调用方使用
- **AND** 函数自身不写出任何文件

#### Scenario: 月份格式校验
- **WHEN** 任意入口传入的 `month` 参数不符合 `YYYYMM` 格式
- **THEN** 处理过程 MUST 以错误终止
- **AND** CLI 以退出码 4 退出
- **AND** JSON 模式下 `error.code` 为 `"processing_error"`

#### Scenario: 日期解析多格式支持
- **WHEN** `出行日期` 列包含 `2026-02-15`、`2026/02/15`、`2026年2月15日` 等不同格式
- **THEN** `parse_date()` 正确解析为 `pd.Timestamp`
- **AND** 用于窗口筛选

## REMOVED Requirements

### Requirement: 工作流 JSON 输出扩展

**Reason**: 此能力的 JSON 输出契约由 `cli-output` capability 统一管理；本次变更后 CLI JSON 输出不再含报表相关字段，`sales-report` capability 不应再单独定义 JSON 扩展字段。该需求的删除避免与 `cli-output` 的新约束冲突。

**Migration**: 见 `cli-output` capability 的 REMOVED 段：调用方不再依赖 `data.report_file` / `data.report_rows`；改为通过退出码与 `data.statistics` 评估处理结果。
