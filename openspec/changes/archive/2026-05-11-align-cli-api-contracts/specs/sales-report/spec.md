## MODIFIED Requirements

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
