## MODIFIED Requirements

### Requirement: Core matching side effects

`process_excel_files()` MUST 保持其公开签名和当前副作用，同时将实现重构为更小的 helper 或移至兼容 facade 之后。It MUST 读取两个输入文件、填充 `支付手续费`，并在返回前调用 `add_sales_report_period()`。因此，即使调用方只请求基础匹配，其返回的 DataFrame 也会包含或刷新 `销售报表账期` 列。

#### Scenario: 基础匹配也会刷新销售报表账期
- **WHEN** `process_excel_files(order_file, payment_file)` 成功完成时
- **THEN** 返回的 DataFrame 包含 `支付手续费` 列
- **AND** 返回的 DataFrame 包含 `销售报表账期` 列
- **AND** `销售报表账期` 已按 `add_sales_report_period()` 重新计算

#### Scenario: 兼容 facade 保持导入路径
- **WHEN** 现有代码从 `utils.py` 导入 `process_excel_files`
- **THEN** 在匹配实现拆分为更小模块后，该导入 SHALL 继续可用
- **AND** 调用该导入函数 SHALL 产生与重构前相同的 DataFrame 语义

### Requirement: 多级匹配优先级

匹配引擎 MUST 在提取内部 helper 时保持当前匹配优先级。It MUST 先尝试 20 字符 exact matching。若没有 exact 候选被采纳，则当前 fallback 扫描会按 payment-file 顺序逐行评估 P-number matching 和 hyphen matching，并采纳首个通过业务类型校验的 P-number OR hyphen 匹配行。这意味着 P-number 当前并不对整份 payment file 中更晚出现的 hyphen 命中拥有全局优先级。

说明：这里的重点不是“P-number 比 hyphen 更强”或相反，而是保留历史实现中可观察到的 payment 行扫描顺序。重构时若将候选预先分组、排序或索引化，必须用 golden tests 证明采纳的 payment 行仍完全一致。

#### Scenario: 20 字符精确匹配优先
- **WHEN** 订单行的 `订单号` 前 20 字符与某条支付记录的 `商户订单号` 前 20 字符相等
- **AND** 业务类型校验通过
- **THEN** 该支付记录被采纳，匹配引擎停止对该订单的后续策略尝试
- **AND** 不再执行 P-number 匹配或连字符匹配

#### Scenario: P-number 回退匹配
- **WHEN** 20 字符精确匹配未命中
- **AND** 从订单行 `外部订单号` 中提取的 `r"P\d+"` 与某条支付记录 `商品名称` 中提取的 `r"P\d+"` 相等
- **AND** 业务类型校验通过
- **THEN** 该支付记录被采纳

#### Scenario: 连字符回退匹配
- **WHEN** 20 字符精确匹配未命中
- **AND** 订单行 `外部订单号` 中的某段（按 `-` 分割）与支付记录 `商品名称` 最后一个 `-` 后的段相等
- **AND** 业务类型校验通过
- **THEN** 在当前 payment 行扫描中该支付记录 MAY be accepted even if a later row would have matched by P-number

#### Scenario: 回退扫描顺序
- **WHEN** exact matching 未采纳任何 payment 行
- **AND** 更早的 payment 行通过 hyphen 且业务类型有效
- **AND** 更晚的 payment 行通过 P-number 且业务类型有效
- **THEN** 采纳更早的 hyphen 匹配，因为 fallback matching 按 payment-file 顺序评估

#### Scenario: helper 提取不改变优先级
- **WHEN** exact 和 fallback matching 通过提取出的 helper functions 实现时
- **THEN** 对于同一组有序输入文件，每个订单可观察到的被采纳 payment 行 SHALL 与重构前实现一致

#### Scenario: 无匹配
- **WHEN** 三种策略均未命中
- **THEN** 该订单的 `支付手续费` 列保持空值（NaN）
- **AND** 在匹配统计中计入未匹配数量

## ADDED Requirements

### Requirement: Matching implementation boundaries

matching implementation SHALL 被拆解为更小、可测试的单元，同时通过 `utils.py` 保持公开兼容性。

#### Scenario: Matching helpers 覆盖业务决策
- **WHEN** matching implementation 被重构时
- **THEN** business-order-column detection、order-amount classification、business-type compatibility、payment-fee extraction、exact matching 和 fallback matching SHALL 能被表达为独立的 helper-level behaviours
- **AND** 这些 helpers SHALL NOT 需要 CLI、HTTP API 或 workflow-service state 即可运行

#### Scenario: Matching module 保持与 adapters 独立
- **WHEN** matching code 执行时
- **THEN** it SHALL NOT 解析 CLI arguments
- **AND** it SHALL NOT 检查 Flask request objects
- **AND** it SHALL NOT 格式化 CLI 或 API JSON responses
