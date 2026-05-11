## Purpose

销售报表能力 - 定义两阶段销售报表工作流：阶段一标注销售报表账期（全退/已取消），阶段二筛选未标注行并按出行日期窗口生成月度报表。该能力由 `utils.py` 中的 `add_sales_report_period()`、`filter_unmarked_and_generate_report()` 与 `process_sales_report_workflow()` 实现。当前 `cli.py` 通过可选位置参数 `target_month` 触发，`excel_merge.py` 通过 `--month` 或交互式月份输入触发。

## Requirements

### Requirement: 销售报表账期标注

`add_sales_report_period()` MUST 写入 `销售报表账期` 列，根据订单状态标注"全退"或"已取消"，未匹配规则的行保持空值。

#### Scenario: 全退标注
- **WHEN** 同一 `订单号` 在数据集中出现多次
- **AND** 这些行的 `订单金额` 之和为 0
- **THEN** 这些行的 `销售报表账期` 列值为 `"全退"`

#### Scenario: 已取消标注
- **WHEN** 订单行的状态字段（如 `订单状态`）字符串包含 `"取消"` 子串
- **AND** 该行 `订单金额` 为 0
- **THEN** 该行 `销售报表账期` 列值为 `"已取消"`

#### Scenario: 全退优先于已取消
- **WHEN** 一行同时满足全退（重复订单号金额合计 0）和已取消（状态含"取消"且金额 0）条件
- **THEN** 标注为 `"全退"`（标注按写入顺序生效）

#### Scenario: 普通订单无标注
- **WHEN** 订单行不满足全退或已取消条件
- **THEN** `销售报表账期` 列保持空值（NaN 或空字符串）

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

`process_sales_report_workflow()` MUST 编排完整流程：调用 `process_excel_files()` 获得已匹配并已标注账期的订单 DataFrame → 计算未标注的报表 DataFrame。工作流函数自身 MUST NOT 写入任何文件；它返回 `(updated_df, filtered_report_df)` 供调用方决定如何持久化。该工作流 MUST 可由 CLI、交互模式与 HTTP API 触发；CLI 与交互模式 MUST 将 `updated_df` 就地写回原始订单文件、不产生其它文件，HTTP API 的文件写入语义由 `http-api` capability 单独定义，本 capability 不约束。

#### Scenario: 完整工作流执行顺序（CLI）
- **WHEN** `cli.py` 收到位置参数 `target_month=202602`
- **THEN** 依次执行：
  1. `process_sales_report_workflow()` 调用 `process_excel_files()` 完成支付手续费匹配，并依赖当前 `process_excel_files()` 内部调用 `add_sales_report_period()` 标注 `销售报表账期` 列
  2. `filter_unmarked_and_generate_report()` 计算筛选 DataFrame（不落盘）并回填 `销售报表YYYYMM`
  3. `write_result_file()` 将更新后的订单 DataFrame 就地写回原始订单文件
- **AND** 任一步骤失败终止后续步骤并向调用方传播异常
- **AND** 不在任何目录生成 `report_YYYYMM.xlsx` 等独立报表文件

#### Scenario: 完整工作流执行顺序（Interactive）
- **WHEN** 用户在交互模式下选择为某个月份生成销售报表
- **THEN** 执行的函数序列与 CLI 工作流相同
- **AND** 同样不产生独立的报表文件

#### Scenario: 完整工作流执行顺序（API）
- **WHEN** HTTP API 在 `/merge` 或 `/merge/json` 路由中以 `month` 参数触发该工作流
- **THEN** 同样调用 `process_sales_report_workflow()`，该函数当前通过 `process_excel_files()` 的内部账期标注副作用完成阶段一标注，然后调用 `filter_unmarked_and_generate_report()`
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
- **AND** `cli.py` 在参数验证阶段以退出码 2 退出
- **AND** `cli.py` JSON 模式下 `error.code` 为 `"usage_error"`

#### Scenario: 日期解析多格式支持
- **WHEN** `出行日期` 列包含 `2026-02-15`、`2026/02/15`、`2026年2月15日` 等不同格式
- **THEN** `parse_date()` 返回 `pd.Timestamp` 或 `None`
- **AND** 中文 `YYYY年M月D日` 形式当前按年月解析并返回该月 1 日
- **AND** `filter_unmarked_and_generate_report()` 当前使用 `pd.to_datetime(..., errors='coerce')` 解析筛选日期，而不是逐行调用 `parse_date()`

### Requirement: Sales-report workflow service invocation

Entry points MUST invoke the full sales-report workflow through the workflow/service layer while preserving the existing sales-report semantics.

#### Scenario: Service delegates to existing sales-report workflow
- **WHEN** the workflow service receives a full sales-report request
- **THEN** it SHALL call the existing sales-report workflow implementation
- **AND** it SHALL preserve the returned updated order DataFrame and filtered report DataFrame

#### Scenario: CLI sales-report persistence through service
- **WHEN** CLI invokes the full sales-report service operation
- **THEN** the service SHALL write the updated order DataFrame back to the original order file
- **AND** it SHALL not persist the filtered report DataFrame as a CLI report file

#### Scenario: API sales-report persistence through service
- **WHEN** API invokes the sales-report service operation
- **THEN** the service SHALL make the filtered report DataFrame available for API result-file persistence
- **AND** it SHALL return API metadata for a downloadable report artifact
