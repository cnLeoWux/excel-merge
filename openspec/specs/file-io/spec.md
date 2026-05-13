## Purpose

文件读写能力 - 定义 Excel/CSV 文件的读取容错策略、写入格式保留、订单号字符串保护与文件查找规则。该能力由 `file_io.py` 承载核心实现，并通过 `utils.py` facade 向现有调用方保持兼容。

## Requirements

### Requirement: 编码自动回退

CSV 文件读取 MUST 按固定顺序尝试多种编码，直到成功或全部失败。回退顺序不得改变（向后兼容）。

该 fallback 顺序属于兼容契约，不只是实现细节。历史文件可能依赖某个较早编码“先成功”的行为，因此不能仅按更常见编码或更现代默认值重新排序。

#### Scenario: GBK 编码优先
- **WHEN** 读取一个 CSV 文件
- **THEN** 首先尝试 `gbk` 编码
- **AND** 仅在 `gbk` 解码失败时才尝试下一种编码

#### Scenario: 编码回退顺序
- **WHEN** `gbk` 解码失败
- **THEN** 依次尝试 `utf-8` → `gb2312` → `latin-1` → `utf-8-sig`
- **AND** 顺序不得调整

#### Scenario: 重构后的 reader 保持 fallback 顺序
- **WHEN** CSV reading 通过 helper functions 或专用 `file_io.py` module 实现时
- **THEN** helper/module SHALL 按相同文档顺序尝试 encodings
- **AND** tests SHALL 能在不调用 CLI 或 HTTP API adapters 的情况下验证 fallback behaviour

#### Scenario: 全部编码失败
- **WHEN** 所有编码均无法成功解码 CSV 文件
- **THEN** 抛出可读异常，包含尝试过的编码列表
- **AND** 不返回静默的空 DataFrame

### Requirement: CSV 分隔符自动检测

读取 CSV 时 MUST 在每种编码下尝试多种分隔符，同时不能静默丢弃异常行。

注释行跳过、分隔符尝试和异常行警告需要作为同一个 reader contract 维护。测试应覆盖“带注释头 + 非逗号分隔 + 中文编码”的组合场景，而不仅是单独 happy path。

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

#### Scenario: CSV helpers 保持 adapter-independent
- **WHEN** CSV reading helpers 被调用时
- **THEN** they SHALL 返回 DataFrames 或抛出可读异常
- **AND** they SHALL NOT 格式化 CLI JSON
- **AND** they SHALL NOT 构造 HTTP responses

### Requirement: Excel 引擎检测

Excel 文件读取 MUST 根据扩展名与文件实际格式选择正确的引擎。

#### Scenario: .xlsx 使用 openpyxl
- **WHEN** 文件扩展名为 `.xlsx`
- **AND** 文件可被识别为 ZIP 容器
- **THEN** 使用 `openpyxl` 引擎读取

#### Scenario: .xlsx 文件实际为 .xls 格式（伪 xlsx）
- **WHEN** 文件扩展名为 `.xlsx`
- **AND** 文件不是 ZIP 容器（`zipfile.BadZipFile`）
- **THEN** 回退到 `xlrd` 引擎读取
- **AND** 不报错退出

#### Scenario: .xls 始终使用 xlrd
- **WHEN** 文件扩展名为 `.xls`
- **THEN** 使用 `xlrd` 引擎读取
- **AND** 不尝试 openpyxl

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

### Requirement: 文件格式保留写入

`write_result_file()` MUST 在实现移至兼容 facade 后，保留原始文件扩展名和写入语义。It MUST 继续可从 `utils.py` 为现有调用方访问。

#### Scenario: CSV 写入保留 CSV
- **WHEN** 原始文件为 `.csv`
- **THEN** 输出文件也为 `.csv`
- **AND** 使用 `utf-8-sig` 编码（兼容 Excel 中文显示）

#### Scenario: Excel 写入保留 .xlsx
- **WHEN** 原始文件路径扩展名为 `.xlsx`
- **THEN** 使用 `openpyxl` 写入同一路径

#### Scenario: .xls 写入尝试保留路径
- **WHEN** 原始文件路径扩展名为 `.xls`
- **THEN** `write_result_file()` SHALL attempt to write to the same `.xls` path
- **AND** it MAY use `xlwt` when available or let pandas choose the writer engine
- **AND** if the environment cannot write `.xls`, the function SHALL propagate the write error to the caller

#### Scenario: 兼容 facade 保持 file I/O imports
- **WHEN** 现有代码从 `utils.py` 导入 `read_file_with_appropriate_method`、`write_result_file` 或 `find_file_path`
- **THEN** 在 file I/O implementation 拆分后，这些 imports SHALL 继续可用
- **AND** 它们可观察到的文件格式与错误语义 SHALL 保持不变

### Requirement: 文件查找路径

`find_file_path()` MUST 按固定顺序在多个目录中查找文件。

#### Scenario: 当前目录优先
- **WHEN** 调用 `find_file_path("order.xlsx")`
- **AND** 当前工作目录存在 `order.xlsx`
- **THEN** 返回当前目录下的路径

#### Scenario: 回退到 ExcelForHandel/
- **WHEN** 当前目录不存在该文件
- **AND** `ExcelForHandel/order.xlsx` 存在
- **THEN** 返回 `ExcelForHandel/` 子目录下的路径

#### Scenario: 文件不存在
- **WHEN** 当前目录和 `ExcelForHandel/` 均无该文件
- **THEN** 返回原始 `Path(filename)`
- **AND** 调用方根据该路径不存在来报告 `file_not_found` 或其它可读错误

### Requirement: File I/O implementation boundaries

File reading、writing、path lookup、identifier-column normalization 以及 CSV/Excel fallback logic SHALL 与 matching、sales-report、CLI 和 HTTP adapter responsibilities 隔离。

#### Scenario: File I/O module 不执行 matching
- **WHEN** file I/O code 读取或写入 DataFrame 时
- **THEN** it SHALL NOT 填充 `支付手续费`
- **AND** it SHALL NOT 计算 `销售报表账期`
- **AND** it SHALL NOT 对 payment business types 分类

#### Scenario: Identifier normalization is centralized
- **WHEN** CSV or Excel readers 加载包含 `订单` 或 `流水` 的列时
- **THEN** identifier string preservation 和 cleanup SHALL 由共享 file I/O normalization logic 处理
- **AND** matching code SHALL 在可能时接收已归一化的 identifier-like values
