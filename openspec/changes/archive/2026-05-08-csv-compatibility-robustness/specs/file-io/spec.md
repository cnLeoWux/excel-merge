## MODIFIED Requirements

### Requirement: 订单号字符串保护

读取阶段 MUST 强制将订单号相关列保留为字符串类型，防止 Excel 数字转换或科学计数法导致前导零丢失。此外，必须处理第三方支付平台由于规避科学计数法而引入的特殊保护符号（如 `="123"` 或制表符）。

#### Scenario: Excel 订单号列保护
- **WHEN** 读取 Excel 文件
- **AND** 列名为 `订单号`、`商户订单号` 或 `商务订单号`
- **THEN** 通过 `dtype={"订单号": str}` 或 `astype(str)` 保留为字符串
- **AND** 长数字订单号不出现 `1.23E+18` 形式

#### Scenario: CSV 订单/流水列保护
- **WHEN** 读取 CSV 文件
- **THEN** 使用 `dtype=str` 读取整个 CSV，防止长数字在加载时就被错误截断或转为浮点数
- **AND** 包含 `"订单"` 或 `"流水"` 子串的列，其值必须被自动清理（剥除两端的制表符、空格、以及 `="` 包装符）

### Requirement: CSV 分隔符自动检测

读取 CSV 时 MUST 在每种编码下尝试多种分隔符，同时不能静默丢弃异常行。

#### Scenario: 分隔符回退顺序
- **WHEN** 某编码读取成功但 DataFrame 列数异常少（放宽至列数 >= 2 即算成功）
- **THEN** 依次尝试分隔符 `,` → `;` → `\t`
- **AND** 最后回退到 `sep=None`（pandas 自动检测）

#### Scenario: CSV 注释行跳过
- **WHEN** CSV 文件首行以 `#` 开头
- **THEN** 跳过所有连续的 `#` 注释行
- **AND** 第一个非注释行作为表头解析

#### Scenario: 异常行处理
- **WHEN** CSV 行中包含多于或少于表头定义的字段数量
- **THEN** 必须输出警告日志 (`on_bad_lines="warn"`)，而不能默默跳过 (`skip`)
