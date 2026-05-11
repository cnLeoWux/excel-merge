## Purpose

核心匹配能力 - 定义订单数据与支付流水的多级匹配算法、业务类型校验规则与金额赋值逻辑。该能力由 `utils.py` 中的 `process_excel_files()` 实现，是工具的核心业务逻辑。

## Requirements

### Requirement: Core matching side effects

`process_excel_files()` MUST currently read both input files, populate `支付手续费`, and then call `add_sales_report_period()` before returning. Therefore its returned DataFrame includes or refreshes the `销售报表账期` column even when the caller requested only basic matching.

#### Scenario: Basic matching also refreshes sales-report period
- **WHEN** `process_excel_files(order_file, payment_file)` completes successfully
- **THEN** the returned DataFrame contains a `支付手续费` column
- **AND** the returned DataFrame contains a `销售报表账期` column
- **AND** `销售报表账期` has been recalculated according to `add_sales_report_period()`

### Requirement: 多级匹配优先级

匹配引擎 MUST first try 20-character exact matching. If no exact candidate is accepted, the current fallback scan evaluates P-number matching and hyphen matching for each payment row in payment-file order, accepting the first row whose P-number OR hyphen match passes business type validation. This means P-number does not currently have global priority over a later hyphen match across the whole payment file.

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

#### Scenario: Fallback scan order
- **WHEN** exact matching accepts no payment row
- **AND** an earlier payment row matches by hyphen with valid business type
- **AND** a later payment row matches by P-number with valid business type
- **THEN** the earlier hyphen match is accepted because fallback matching is evaluated in payment-file order

#### Scenario: 无匹配
- **WHEN** 三种策略均未命中
- **THEN** 该订单的 `支付手续费` 列保持空值（NaN）
- **AND** 在匹配统计中计入未匹配数量

### Requirement: 业务类型校验

所有匹配策略 MUST 在采纳支付记录前执行业务类型校验。订单方向（正单/退单）必须与支付记录类型一致，否则视为不匹配。

#### Scenario: 正单匹配收费/服务费
- **WHEN** 订单 `订单金额` > 0
- **AND** 候选支付记录的业务类型字段为"收费"或"服务费"
- **THEN** 业务类型校验通过

#### Scenario: 退单匹配退费/退款
- **WHEN** 订单 `订单金额` < 0
- **AND** 候选支付记录的业务类型字段为"退费"或"退款"
- **THEN** 业务类型校验通过

#### Scenario: 业务类型不一致拒绝采纳
- **WHEN** 正单（金额 > 0）匹配到一条业务类型为"退费"的支付记录
- **THEN** 该候选被拒绝
- **AND** 匹配引擎继续尝试其他候选或后续策略

### Requirement: 金额赋值规则

`支付手续费` 列的赋值 MUST 根据订单方向选择支付记录的对应金额字段。

#### Scenario: 正单赋值支出金额
- **WHEN** 订单为正单（金额 > 0）且匹配成功
- **THEN** `支付手续费` = 支付记录的 `支出金额（-元）`（预期为负值）

#### Scenario: 退单赋值收入金额
- **WHEN** 订单为退单（金额 < 0）且匹配成功
- **THEN** `支付手续费` = 支付记录的 `收入金额（+元）`（预期为正值）

#### Scenario: 零金额订单短路
- **WHEN** 订单 `订单金额` 为 0 或缺失
- **THEN** `支付手续费` 直接赋值为 `0.0`
- **AND** 不执行任何匹配策略
- **AND** 不消耗候选支付记录

### Requirement: P-number 提取规则

P-number 提取 MUST 使用正则 `r"P\d+"`，区分大小写，从字符串中匹配第一个出现的模式。

#### Scenario: 标准 P-number 提取
- **WHEN** 输入字符串为 `"P2507021103060001-extra"`
- **THEN** `extract_p_number()` 返回 `"P2507021103060001"`

#### Scenario: 小写 p 不匹配
- **WHEN** 输入字符串为 `"p2507021103060001"`
- **THEN** `extract_p_number()` 返回 `None`

#### Scenario: 无 P-number
- **WHEN** 输入字符串不包含 `P` 后接数字的模式
- **THEN** `extract_p_number()` 返回 `None`

### Requirement: 列名灵活匹配

匹配引擎 MUST 通过子串搜索定位业务订单号列，以容忍上游数据源的列名差异。

#### Scenario: 商户订单号列识别
- **WHEN** 支付文件存在列名同时包含 `"商户"` 和 `"订单"` 子串
- **THEN** 该列被选为业务订单号列
- **AND** 用于 20 字符精确匹配

#### Scenario: 列名识别回退
- **WHEN** 不存在同时包含 `"商户"` 和 `"订单"` 的列
- **AND** 存在包含 `"订单"` 子串的列
- **THEN** 第一个包含 `"订单"` 的列被选为业务订单号列
