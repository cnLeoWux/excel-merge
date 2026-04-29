# Excel Merge Tool - Architecture

## Overview

Excel Merge Tool 是一个订单数据与支付流水自动匹配工具。核心功能是将订单 Excel/CSV 文件与支付/退费文件进行匹配，填充"支付手续费"列。同时支持销售报表账期标记和月度报表生成。

技术栈：Python 3.7+、pandas、openpyxl、xlrd、Flask。

---

## System Architecture

### Entry Points

系统提供 4 种入口方式，均调用 `utils.py` 中的核心逻辑：

| 入口 | 文件 | 说明 | 控制台命令 |
|------|------|------|-----------|
| 交互式 | `excel_merge.py` | 从 `ExcelForHandel/` 列出文件供用户选择 | `excel-merge` |
| CLI | `cli.py` | argparse 参数模式，支持 `--month`；结果就地写回订单文件 | `excel-merge-cli` |
| Flask API | `excel_merge_api.py` | HTTP 服务，支持文件上传和下载 | 无（直接运行） |
| 控制台脚本 | `setup.py` | 通过 `pip install -e .` 注册的 entry_points | — |

### Dependency Graph

```
excel_merge.py ──┐
cli.py ──────────┤──→ utils.py (全部核心逻辑, ~930 行)
excel_merge_api.py┘        │
                           ↓
                   pandas, openpyxl, xlrd, flask
```

所有业务逻辑集中在 `utils.py` 单一模块中。3 个入口脚本仅负责用户交互和参数解析。

---

## Core Module: utils.py

### Public Functions

| 函数 | 行号 | 职责 |
|------|------|------|
| `extract_p_number(text)` | L15 | 从字符串中提取 `r"P\d+"` 模式 |
| `match_orders_by_p_number(ext_no, prod_name)` | L27 | 比较两个字段的 P-number 是否一致 |
| `read_file_with_appropriate_method(file_path)` | L39 | CSV/Excel 统一读取，含编码回退链 |
| `process_excel_files(order, payment, verbose)` | L189 | 主匹配循环：精确→P-number→连字符 |
| `find_file_path(filename)` | L520 | 搜索当前目录和 `ExcelForHandel/` |
| `write_result_file(df, file_path)` | L539 | 写入结果，保持原文件格式 |
| `add_sales_report_period(order_df, verbose)` | L572 | 标记销售报表账期（全退、已取消） |
| `parse_date(date_val)` | L687 | 多格式日期解析器 |
| `get_year_month(date_val)` | L726 | 日期 → "YYYYMM" 字符串 |
| `filter_unmarked_and_generate_report(...)` | L743 | 筛选未标记数据，生成月度报表 |
| `process_sales_report_workflow(...)` | L887 | 编排完整销售报表工作流 |

---

## Data Flow

### Basic Matching Flow

```
输入文件 (订单 + 支付)
    │
    ▼
read_file_with_appropriate_method()  ← 编码回退 / 引擎检测
    │
    ▼
process_excel_files()                ← 三级匹配 + 业务类型校验
    │
    ▼
write_result_file()                  ← 保持 CSV/Excel 原格式
    │
    ▼
输出文件（支付手续费已填充）
```

### Sales Report Flow (--month)

```
process_sales_report_workflow()
    │
    ├──→ process_excel_files()           ← 步骤1: 匹配支付手续费
    │
    ├──→ add_sales_report_period()       ← 步骤2: 标记全退/已取消
    │       ├── 重复订单号金额合计=0 → "全退"
    │       └── 状态含"取消"且金额=0 → "已取消"
    │
    └──→ filter_unmarked_and_generate_report()  ← 步骤3: 内存中筛选并标记
            ├── 过滤已标记行
            ├── 筛选出行日期在目标月份前1年范围内的数据
            ├── 在原 DataFrame 中将这些行标记为"销售报表YYYYMM"
            └── 返回 (updated_df, report_df)，由调用方就地写回订单文件
                （不生成独立的 report_YYYYMM.xlsx 文件）
```

---

## Matching Algorithm

三级优先匹配，每次匹配均需通过业务类型校验：

### 1. 精确匹配（20字符截断）
- 订单文件 `订单号` 前20字符 ↔ 支付文件 `商户订单号` 前20字符
- 商户订单号列通过子字符串搜索定位：`"商户" in col and "订单" in col`

### 2. P-number 匹配
- 正则 `r"P\d+"`（区分大小写）分别从 `外部订单号` 和 `商品名称` 中提取
- 两个 P-number 相等则匹配

### 3. 连字符匹配
- `外部订单号` ↔ `商品名称` 最后一个 `-` 之后的部分

### Business Type Gate

所有匹配必须通过业务类型校验：

| 订单金额 | 订单类型 | 允许的业务类型 | 取值列 |
|----------|---------|--------------|--------|
| > 0 | 正单 | 收费、服务费 | `支出金额（-元）` |
| < 0 | 退单 | 退费、退款 | `收入金额（+元）` |
| = 0 | 跳过 | — | 直接设为 0.0 |

---

## File I/O

### Encoding Fallback Chain (CSV)
```
gbk → utf-8 → gb2312 → latin-1 → utf-8-sig
```
每种编码依次尝试分隔符 `,`、`;`、`\t`，最后用 `sep=None` 自动检测。

### Excel Engine Detection
- `.xlsx`：zipfile 探测 → 成功用 openpyxl，`BadZipFile` 则用 xlrd
- `.xls`：始终用 xlrd

### CSV 特殊处理
- `#` 开头的行视为注释跳过，第一个非注释行为表头
- 读取后包含 `"订单"` 或 `"流水"` 的列强制转为 `str` 类型
- 写入时使用 `utf-8-sig` 编码

---

## Flask API Endpoints

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | Web 测试页面 |
| GET | `/health` | 健康检查 |
| POST | `/merge` | 上传文件，直接返回处理后的文件 |
| POST | `/merge/json` | 上传文件，返回 JSON（含下载链接和统计） |
| GET | `/download/<filename>` | 下载结果文件 |

服务运行在 `0.0.0.0:5000`，默认 debug 模式。上传文件暂存 `uploads/`，结果存 `results/`。

---

## Known Limitations

- **性能**：`iterrows()` + 嵌套循环 → O(n*m) 复杂度，大文件处理慢
- **异常处理**：多处裸 `except Exception` 可能隐藏错误
- **Magic Number**：`[:20]` 截断硬编码，无常量定义
- **类型转换**：`astype(str)` 会将 NaN 转为字面量 `"nan"`
- **日志**：导入了 `logging` 模块但实际使用 `print` 输出
- **API**：`/merge` 固定返回 XLSX mimetype，不论实际文件格式
- **就地覆盖**：CLI 始终覆盖原订单文件（已无 `-o`/`--output-dir`），调用前需自行备份
