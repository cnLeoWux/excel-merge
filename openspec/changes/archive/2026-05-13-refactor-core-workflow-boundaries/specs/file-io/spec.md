## MODIFIED Requirements

### Requirement: 编码自动回退

CSV 文件读取 MUST 即使实现从 `utils.py` 移至专用 file I/O module，也保持固定的 encoding fallback 顺序。该 fallback 顺序 MUST 保持向后兼容。

说明：fallback 顺序属于兼容契约，不只是实现细节。历史文件可能依赖某个较早编码“先成功”的行为，因此不能仅按更常见编码或更现代默认值重新排序。

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

读取 CSV 时 MUST 在允许实现拆分为专用 helper functions 的同时，保留 separator fallback behaviour。It MUST 在每种 encoding 下尝试支持的 separators，并且 must not 静默丢弃格式异常的行。

说明：注释行跳过、分隔符尝试和异常行警告需要作为同一个 reader contract 维护。测试应覆盖“带注释头 + 非逗号分隔 + 中文编码”的组合场景，而不仅是单独 happy path。

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

## ADDED Requirements

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
