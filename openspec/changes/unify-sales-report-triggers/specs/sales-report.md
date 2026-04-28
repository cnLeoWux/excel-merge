## MODIFIED Requirements

### Requirement: 端到端工作流编排

`process_sales_report_workflow()` MUST 编排完整两阶段流程：匹配支付手续费 → 标注账期 → 筛选并生成报表。This workflow MUST be triggerable from the CLI, interactive mode, and the HTTP API.

#### Scenario: 完整工作流执行顺序（CLI）
- **WHEN** CLI 收到 `--month 202602` 参数
- **THEN** 依次执行：
  1. `process_excel_files()` 完成支付手续费匹配
  2. `add_sales_report_period()` 标注账期列
  3. 写入更新后的订单文件（默认原地，或 `-o` 指定路径）
  4. `filter_unmarked_and_generate_report()` 生成 `report_YYYYMM.xlsx`
- **AND** 任一步骤失败不得跳过后续步骤的错误处理

#### Scenario: 完整工作流执行顺序（Interactive）
- **WHEN** a user opts to generate a sales report for a given month in interactive mode
- **THEN** the same sequence of functions as the CLI workflow SHALL be executed.

#### Scenario: 完整工作流执行顺序（API）
- **WHEN** an API call to `/merge` or `/merge/json` includes the `month` parameter
- **THEN** the same sequence of functions as the CLI workflow SHALL be executed.

#### Scenario: 月份格式校验
- **WHEN** the `month` parameter from any entry point does not conform to `YYYYMM` format
- **THEN** the process MUST terminate with an error.
- **AND** the CLI SHALL exit with code 4.
- **AND** the API SHALL return a 4xx error with a descriptive message.

#### Scenario: 日期解析多格式支持
- **WHEN** `出行日期` 列包含 `2026-02-15`、`2026/02/15`、`2026年2月15日` 等不同格式
- **THEN** `parse_date()` 正确解析为 `pd.Timestamp`
- **AND** 用于窗口筛选
