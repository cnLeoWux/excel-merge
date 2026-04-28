## Purpose

文件读写能力 - 定义 Excel/CSV 文件的读取容错策略、写入格式保留、订单号字符串保护与文件查找规则。该能力由 `utils.py` 中的 `read_file_with_appropriate_method()`、`write_result_file()` 与 `find_file_path()` 实现。

## Requirements

### Requirement: 编码自动回退

CSV 文件读取 MUST 按固定顺序尝试多种编码，直到成功或全部失败。回退顺序不得改变（向后兼容）。

#### Scenario: GBK 编码优先
- **WHEN** 读取一个 CSV 文件
- **THEN** 首先尝试 `gbk` 编码
- **AND** 仅在 `gbk` 解码失败时才尝试下一种编码

#### Scenario: 编码回退顺序
- **WHEN** `gbk` 解码失败
- **THEN** 依次尝试 `utf-8` → `gb2312` → `latin-1` → `utf-8-sig`
- **AND** 顺序不得调整

#### Scenario: 全部编码失败
- **WHEN** 所有编码均无法成功解码 CSV 文件
- **THEN** 抛出可读异常，包含尝试过的编码列表
- **AND** 不返回静默的空 DataFrame

### Requirement: CSV 分隔符自动检测

读取 CSV 时 MUST 在每种编码下尝试多种分隔符。

#### Scenario: 分隔符回退顺序
- **WHEN** 某编码读取成功但 DataFrame 列数异常少（启发式判定）
- **THEN** 依次尝试分隔符 `,` → `;` → `\t`
- **AND** 最后回退到 `sep=None`（pandas 自动检测）

#### Scenario: CSV 注释行跳过
- **WHEN** CSV 文件首行以 `#` 开头
- **THEN** 跳过所有连续的 `#` 注释行
- **AND** 第一个非注释行作为表头解析

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

读取阶段 MUST 强制将订单号相关列保留为字符串类型，防止 Excel 数字转换或科学计数法导致前导零丢失。

#### Scenario: Excel 订单号列保护
- **WHEN** 读取 Excel 文件
- **AND** 列名为 `订单号` 或 `外部订单号`
- **THEN** 通过 `dtype={"订单号": str}` 或 `astype(str)` 保留为字符串
- **AND** 长数字订单号不出现 `1.23E+18` 形式

#### Scenario: CSV 订单/流水列保护
- **WHEN** 读取 CSV 文件
- **AND** 列名包含 `"订单"` 或 `"流水"` 子串
- **THEN** 该列在读取后转换为字符串类型

### Requirement: 文件格式保留写入

`write_result_file()` MUST 根据原始文件扩展名选择写入格式，不得跨格式转换。

#### Scenario: CSV 写入保留 CSV
- **WHEN** 原始文件为 `.csv`
- **THEN** 输出文件也为 `.csv`
- **AND** 使用 `utf-8-sig` 编码（兼容 Excel 中文显示）

#### Scenario: Excel 写入保留 .xlsx
- **WHEN** 原始文件为 `.xlsx` 或 `.xls`
- **THEN** 输出文件为 `.xlsx`（统一使用 openpyxl 写入）

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
- **THEN** 返回 `None`（或调用方据此抛出 `file_not_found` 错误）
