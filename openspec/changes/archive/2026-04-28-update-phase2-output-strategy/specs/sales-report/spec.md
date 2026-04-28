## MODIFIED Requirements

### Requirement: 端到端工作流编排

`process_sales_report_workflow()` MUST 编排完整两阶段流程：匹配支付手续费 → 标注账期 → 筛选并生成报表。系统 SHALL 确保 Phase 2 的报表生成逻辑与 Phase 1 的原文件保存逻辑解耦，即使原文件保存失败，报表也应当被生成。

#### Scenario: 完整工作流执行顺序
- **WHEN** CLI 收到 `--month 202602` 参数
- **THEN** 依次执行：
  1. `process_excel_files()` 完成支付手续费匹配
  2. `add_sales_report_period()` 标注账期列
  3. `filter_unmarked_and_generate_report()` 生成 `report_YYYYMM.xlsx`
  4. 写入更新后的订单文件（默认原地，或 `-o` 指定路径）
- **AND** 第 4 步（写入订单文件）失败不应当影响第 3 步的执行或结果汇报

#### Scenario: 原文件锁定不影响报表生成
- **WHEN** 原始订单文件被其他程序锁定导致写入失败
- **THEN** 系统 MUST 捕获该异常
- **AND** 仍然完成 `report_YYYYMM.xlsx` 的生成与保存
- **AND** 在 CLI 输出中报告报表生成成功，同时提示原文件更新失败

#### Scenario: 月份格式校验
- **WHEN** `--month` 参数不符合 `YYYYMM` 格式（如 `2026-02` 或 `202613`）
- **THEN** CLI 以退出码 4 终止
- **AND** JSON 输出 `error.code` 为 `"processing_error"`

#### Scenario: 日期解析多格式支持
- **WHEN** `出行日期` 列包含 `2026-02-15`、`2026/02/15`、`2026年2月15日` 等不同格式
- **THEN** `parse_date()` 正确解析为 `pd.Timestamp`
- **AND** 用于窗口筛选

### Requirement: 工作流 JSON 输出扩展

当 `--month` 触发销售报表工作流时，CLI 的 JSON 信封 `data` 字段 MUST 在基础统计之外额外包含报表相关字段，并支持报告部分成功的警告信息。

#### Scenario: 报表字段在 JSON 中暴露
- **WHEN** `python cli.py order.xlsx payment.xlsx --month 202602 --json` 成功执行
- **THEN** `data` 包含 `report_file`（字符串路径，或 `null`）
- **AND** `data` 包含 `report_rows`（整数，未生成报表时为 `0`）
- **AND** 基础字段 `output_file` 与 `statistics` 仍存在

#### Scenario: 部分成功时的警告信息
- **WHEN** 报表生成成功但原文件保存失败
- **THEN** JSON 信封的 `ok` 仍可为 `true`（只要核心目标报表已生成）
- **AND** `data` MUST 包含 `warnings` 数组，其中包含描述原文件保存失败的消息
- **AND** `output_file` 应反映预期的路径，即使未成功写入
