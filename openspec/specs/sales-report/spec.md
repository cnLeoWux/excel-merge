## Purpose

销售报表能力 - 定义两阶段销售报表工作流：阶段一标注销售报表账期（全退/已取消），阶段二筛选未标注行并按出行日期窗口生成月度报表。该能力由 `utils.py` 中的 `add_sales_report_period()`、`filter_unmarked_and_generate_report()` 与 `process_sales_report_workflow()` 实现，通过 CLI `--month YYYYMM` 触发。

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

`filter_unmarked_and_generate_report()` MUST 仅保留未标注且出行日期落在目标月份前后 1 年窗口内的行，并写入独立报表文件。

#### Scenario: 仅保留未标注行
- **WHEN** 生成月度报表
- **THEN** 所有 `销售报表账期` 为 `"全退"` 或 `"已取消"` 的行被排除
- **AND** 仅未标注（NaN 或空）的行进入报表

#### Scenario: 出行日期窗口
- **WHEN** 目标月份为 `202602`（2026 年 2 月）
- **THEN** 报表仅保留 `出行日期` ∈ [2025-02-01, 2027-02-28] 的行（目标月份前后 1 年）
- **AND** `出行日期` 缺失或无法解析的行被排除

#### Scenario: 报表文件命名
- **WHEN** 目标月份为 `YYYYMM`
- **THEN** 输出文件名为 `report_YYYYMM.xlsx`
- **AND** 写入到 `--output-dir` 指定目录（默认当前工作目录）

#### Scenario: 空结果不生成空文件
- **WHEN** 筛选后无任何符合条件的行
- **THEN** 不生成报表文件
- **AND** 在 JSON 输出中 `report_file` 为 `null`，`report_rows` 为 `0`

### Requirement: 端到端工作流编排

`process_sales_report_workflow()` MUST 编排完整两阶段流程：匹配支付手续费 → 标注账期 → 筛选并生成报表。

#### Scenario: 完整工作流执行顺序
- **WHEN** CLI 收到 `--month 202602` 参数
- **THEN** 依次执行：
  1. `process_excel_files()` 完成支付手续费匹配
  2. `add_sales_report_period()` 标注账期列
  3. 写入更新后的订单文件（默认原地，或 `-o` 指定路径）
  4. `filter_unmarked_and_generate_report()` 生成 `report_YYYYMM.xlsx`
- **AND** 任一步骤失败不得跳过后续步骤的错误处理

#### Scenario: 月份格式校验
- **WHEN** `--month` 参数不符合 `YYYYMM` 格式（如 `2026-02` 或 `202613`）
- **THEN** CLI 以退出码 4 终止
- **AND** JSON 输出 `error.code` 为 `"processing_error"`

#### Scenario: 日期解析多格式支持
- **WHEN** `出行日期` 列包含 `2026-02-15`、`2026/02/15`、`2026年2月15日` 等不同格式
- **THEN** `parse_date()` 正确解析为 `pd.Timestamp`
- **AND** 用于窗口筛选

### Requirement: 工作流 JSON 输出扩展

当 `--month` 触发销售报表工作流时，CLI 的 JSON 信封 `data` 字段 MUST 在基础统计之外额外包含报表相关字段。

#### Scenario: 报表字段在 JSON 中暴露
- **WHEN** `python cli.py order.xlsx payment.xlsx --month 202602 --json` 成功执行
- **THEN** `data` 包含 `report_file`（字符串路径，或 `null`）
- **AND** `data` 包含 `report_rows`（整数，未生成报表时为 `0`）
- **AND** 基础字段 `output_file` 与 `statistics` 仍存在
