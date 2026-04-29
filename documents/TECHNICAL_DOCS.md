# Excel Merge Tool - Technical Documentation

## Overview

本文档详细描述 Excel Merge Tool 的内部实现细节，包括文件读取管道、匹配算法内部逻辑、销售报表处理、日期解析、列检测和边界情况处理。

所有核心逻辑集中在 `utils.py`（约930行），3个入口脚本（`excel_merge.py`、`cli.py`、`excel_merge_api.py`）为薄包装层。

---

## File Reading Pipeline

### read_file_with_appropriate_method() (L39-186)

统一的文件读取入口，根据文件扩展名分发到 CSV 或 Excel 读取逻辑。

### CSV 读取流程

**编码回退链**（按顺序尝试）：
```
gbk → utf-8 → gb2312 → latin-1 → utf-8-sig
```

**分隔符重试**：每种编码下依次尝试 `,`、`;`、`\t`，最后用 `sep=None`（pandas 自动检测）。

**注释行处理**：
1. 使用 `readlines()` 读取整个文件计算 `#` 开头的行数
2. 将该行数传给 `pd.read_csv()` 的 `skiprows` 参数
3. 注意：`readlines()` 会将整个文件加载到内存

**读取成功判定**：`df.shape[1] > 5` 作为启发式条件判断列数是否足够（阈值 5 为硬编码）。

**类型强制转换**：
- 列名包含 `"订单"` 或 `"流水"` 的列 → `astype(str)`
- 防止 pandas 将长数字订单号解析为浮点数
- 副作用：NaN 值会被转为字面量字符串 `"nan"`

**容错机制**：
- `on_bad_lines="skip"` 静默跳过格式错误的行
- 使用 `python` 引擎作为 C 引擎失败时的回退

### Excel 读取流程

**引擎检测**（.xlsx 文件）：
```python
try:
    zipfile.ZipFile(path, "r")  # 探测是否为合法 ZIP
    engine = "openpyxl"
except BadZipFile:
    engine = "xlrd"             # 可能是伪装的 .xls
```

**引擎选择**：
- `.xlsx` → zipfile 探测后选 openpyxl 或 xlrd
- `.xls` → 始终使用 xlrd

**订单号列类型**：通过 `dtype={"订单号": str}` 读取时直接指定字符串类型。

---

## Column Detection

### 商户订单号列定位

采用子字符串搜索策略：
```python
# 优先匹配
col where ("商户" in col and "订单" in col)

# 回退匹配
col where ("订单" in col)
```

遍历支付文件的所有列名，找到第一个同时包含"商户"和"订单"的列作为商户订单号列。若未找到，则回退到任何包含"订单"的列。

### 订单号列强制字符串

读取后对包含 `"订单"` 或 `"流水"` 关键字的列执行 `astype(str)`，防止长数字被解析为浮点数（如 `4.025070e+21`）。

---

## Matching Algorithm Internals

### process_excel_files() (L189-517)

**迭代方式**：外层 `order_df.iterrows()`，内层遍历 `payment_df`，复杂度 O(n*m)。

### 匹配流程（每行订单）

```
1. 获取订单号，截取前20字符
2. 获取外部订单号
3. 解析订单金额 → 判断正单/退单/零金额
4. 零金额 → 支付手续费=0.0, 跳过匹配
5. 遍历支付记录：
   a. 精确匹配：订单号[:20] == 商户订单号[:20]
   b. P-number匹配：extract_p_number(外部订单号) == extract_p_number(商品名称)
   c. 连字符匹配：外部订单号 == 商品名称.rsplit("-", 1)[-1]
   d. 业务类型校验：正单→收费/服务费，退单→退费/退款
   e. 全部条件满足 → 赋值并跳出内层循环
```

### 业务类型校验

| 条件 | 允许的业务类型 | 取值列 |
|------|--------------|--------|
| `订单金额 > 0`（正单） | `"收费"` 或 `"服务费"` | `支出金额（-元）`（预期负值） |
| `订单金额 < 0`（退单） | `"退费"` 或 `"退款"` | `收入金额（+元）`（预期正值） |
| `订单金额 == 0` | 不匹配 | 直接设为 `0.0` |

注意：金额列名使用全角括号 `（` `）`，必须精确匹配。

### P-number 提取

```python
re.search(r"P\d+", text_str)
```
- 区分大小写：不匹配小写 `p`
- 无分隔符要求：`P2507021103060001` 整体提取
- 输入为 NaN 或 None 时返回 None

---

## Sales Report Workflow

### Phase 1: add_sales_report_period() (L572-684)

为订单数据添加"销售报表账期"列，标记两类特殊订单：

**全退标记**：
1. 按订单号分组
2. 找出重复订单号（出现2次以上）
3. 对每组计算订单金额合计
4. 合计为0 → 标记该组所有行为"全退"
5. 合计不为0 → 不处理

**已取消标记**：
1. 检查"订单状态"列（通过子字符串匹配 `"状态" in col` 定位）
2. 状态包含"取消"且订单金额为0 → 标记为"已取消"
3. 金额不为0 → 不处理

### Phase 2: filter_unmarked_and_generate_report() (L743+)

**筛选逻辑**：
1. 过滤掉"销售报表账期"列已有值的行
2. 解析"出行日期"列
3. 计算时间窗口：目标月份往前推1年（如 `202602` → `2025-02-01` 至 `2026-02-28`）
4. 筛选出行日期在窗口内的行
5. 在原 DataFrame 中将这些行的"销售报表账期"列回填为"销售报表YYYYMM"
6. 返回 `(updated_df, report_df)` 元组，**不**再写入 `report_YYYYMM.xlsx` 文件；调用方（`cli.py` / `excel_merge_api.py`）负责持久化

### Orchestration: process_sales_report_workflow() (L887-928)

编排完整流程：
```
process_excel_files()              → 匹配支付手续费
add_sales_report_period()          → 标记全退/已取消（内含于 filter 中调用）
filter_unmarked_and_generate_report() → 筛选 + 生成报表
```

由 `cli.py` 的 `--month` 参数触发。

---

## Date Parsing

### parse_date() (L687-723)

多格式日期解析器，按优先级尝试：

1. 已经是 `pd.Timestamp` → 直接返回
2. `datetime` 对象 → 转为 `pd.Timestamp`
3. 字符串 → `pd.to_datetime()` 自动解析
4. 中文日期格式 → 正则 `r"(\d{4})[年](\d{1,2})[月](\d{0,2})"` 提取年月

解析失败返回 `None`。

### get_year_month() (L726-740)

日期值 → `"YYYYMM"` 字符串，内部调用 `parse_date()` 后用 `strftime("%Y%m")` 格式化。

---

## File Writing

### write_result_file() (L539-569)

保持原文件格式写入：
- CSV → `df.to_csv(encoding="utf-8-sig")`
- Excel → 根据扩展名和 zipfile 探测选择引擎（与读取逻辑一致）

注意：写入 Excel 时对已存在的 `.xlsx` 文件再次进行 zipfile 探测来选择引擎，这在文件已被 `process_excel_files` 处理后可能不必要。

---

## Performance Characteristics

| 维度 | 说明 |
|------|------|
| 时间复杂度 | O(n*m)，n=订单数，m=支付记录数 |
| 空间复杂度 | 整个文件加载为 DataFrame |
| CSV 注释处理 | `readlines()` 额外读取整个文件 |
| 瓶颈 | `iterrows()` 逐行迭代，无向量化操作 |

---

## Known Issues

### 异常处理
- 多处裸 `except Exception`（L104, L140, L161, L562, L706, L713）捕获所有异常，包括 `SystemExit` 和 `KeyboardInterrupt`
- 可能隐藏底层错误，增加调试难度

### 类型安全
- `astype(str)` 将 NaN 转为字面量 `"nan"`，后续字符串比较可能产生误匹配
- 订单金额转换失败时静默设为0，可能导致错误的零金额处理

### Magic Numbers
- `[:20]` 订单号截断长度硬编码，无命名常量
- `df.shape[1] > 5` 作为读取成功的列数阈值

### 日志
- 导入了 `logging` 模块但未使用，实际通过 `print` 输出
- `verbose` 参数控制输出，但无日志级别概念

### API
- `/merge` 端点固定返回 XLSX mimetype，不论实际输出格式（CSV 文件也返回 XLSX mimetype）
- `MAX_CONTENT_LENGTH` 声明但未通过 `app.config` 生效
- `uploads/` 和 `results/` 目录在模块导入时创建
