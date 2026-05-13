# Design: 核心工作流边界重构

## 设计原则

本变更是行为保持型重构。所有拆分都必须服从已有业务契约：

- 匹配顺序保持：20 字符 exact 优先；exact 未命中后按 payment 文件行顺序扫描 fallback，在每一行上评估 P-number 或 hyphen。
- 业务类型校验保持：正单匹配收费/服务费，退单匹配退费/退款，零金额短路为 `支付手续费 = 0.0`。
- P-number 保持 `r"P\d+"` 且大小写敏感。
- CSV 编码 fallback 保持 `gbk → utf-8 → gb2312 → latin-1 → utf-8-sig`。
- CSV 分隔符 fallback 保持 `, → ; → \t → sep=None`。
- `process_excel_files()` 继续刷新 `销售报表账期`。
- CLI/API 输出 shape 不改变。

## 目标结构

```text
excel-merge/
├── cli.py                  # CLI adapter：参数、输出、退出码
├── excel_merge.py          # 交互 adapter
├── excel_merge_api.py      # HTTP adapter
├── workflow_service.py     # 应用编排、统计、持久化协调、错误归一化
├── utils.py                # 兼容 facade：旧函数名导入/转发
├── file_io.py              # 文件读取/写入/查找
├── matching.py             # 支付手续费匹配核心
└── sales_report.py         # 销售报表账期与筛选
```

职责流向：

```text
Adapters
  ├─ cli.py
  ├─ excel_merge.py
  └─ excel_merge_api.py
        │
        ▼
workflow_service.py
        │
        ├─────────────┬──────────────┬───────────────┐
        ▼             ▼              ▼               ▼
   file_io.py     matching.py   sales_report.py    utils.py facade
```

`utils.py` 需要继续支持旧导入路径：

```text
from utils import process_excel_files
from utils import read_file_with_appropriate_method
```

但其内部可以转发到：

```text
from matching import process_excel_files
from file_io import read_file_with_appropriate_method
```

## 模块边界

### `file_io.py`

负责：

- `read_file_with_appropriate_method(file_path)`
- `write_result_file(df, file_path)`
- `find_file_path(filename)`
- CSV 注释行跳过
- CSV 编码与分隔符 fallback
- Excel 引擎检测
- 订单号/流水号相关列字符串保护和清理

建议内部 helper：

- `_read_csv_with_fallback(path, verbose=False)`
- `_read_excel_with_engine_detection(path, verbose=False)`
- `_count_leading_comment_lines(path, encoding)`
- `_normalize_identifier_columns(df)`
- `_clean_identifier_value(value)`

所有 helper 若暴露 `verbose`，应遵循 `verbose: bool = False` 模式。

边界说明：`file_io.py` 可以做“让 DataFrame 可被业务层安全消费”的清理，例如订单号/流水号列的字符串保护；但不应根据业务含义填充 `支付手续费`、`销售报表账期` 或匹配统计。这样可以避免读取同一文件时因调用场景不同而产生不同业务结果。

### `matching.py`

负责：

- `extract_p_number(text)`
- `match_orders_by_p_number(ext_no, prod_name)`
- `process_excel_files(order_file, payment_file, verbose=False)`
- 订单金额分类
- 支付业务类型校验
- 支付手续费金额提取
- exact/fallback 候选查找

建议内部 helper：

- `_detect_business_order_column(payment_df)`
- `_classify_order_amount(amount)` 返回 `positive` / `negative` / `zero`
- `_is_business_type_compatible(order_direction, business_type)`
- `_extract_payment_fee(payment_row, order_direction)`
- `_matches_exact_order(order_no, merchant_order_no)`
- `_matches_hyphen_fallback(external_order_no, product_name)`
- `_find_exact_match(order_row, payment_df, business_order_col)`
- `_find_fallback_match(order_row, payment_df)`

关键约束：

```text
exact phase:
  scan payment rows until exact + business type accepted

fallback phase:
  for each payment row in original order:
      if P-number matches and business type accepted: accept
      else if hyphen matches and business type accepted: accept
```

不能改成：

```text
try all P-number matches globally
then try all hyphen matches globally
```

因为这会改变“较早 hyphen 命中优先于较晚 P-number 命中”的现有行为。

实现提示：如果引入 helper，fallback helper 应返回“当前行是否命中/命中的手续费”，而不是先收集所有 P-number 候选再排序。否则代码看起来更清晰，但会破坏历史数据依赖的行顺序语义。

### `sales_report.py`

负责：

- `add_sales_report_period(order_df, verbose=False)`
- `parse_date(date_val)`
- `get_year_month(date_val)`
- `filter_unmarked_and_generate_report(updated_df, target_month, verbose=False)`
- `process_sales_report_workflow(order_file, payment_file, target_month, verbose=False)`

`process_sales_report_workflow()` 可以调用 `matching.process_excel_files()`，并依赖其当前账期刷新副作用。若未来希望移除此副作用，应另开行为变更，而不是包含在本重构中。

注意：销售报表模块可以复用 file I/O 的读取/写入能力，但不要反向依赖 workflow service。core module 应保持可在单元测试中直接调用，便于锁定日期解析、全退/已取消标注和筛选窗口。

### `workflow_service.py`

负责：

- 校验文件存在、月份格式、模式选择。
- 调用 core 模块。
- 调用 `write_result_file()` 协调 CLI/interactive/API 持久化。
- 计算共享统计。
- 将已知失败归一化为 `WorkflowError`。

不负责：

- 编码检测细节。
- 匹配算法细节。
- Flask request 或 argparse 细节。
- CLI JSON envelope 或 API JSON response shape。

边界说明：service 可以返回结构化结果和归一化错误码，但 stdout/stderr、Flask response、下载 URL 的具体 JSON shape 仍属于 adapter。这样 CLI 与 HTTP API 可以继续保持各自历史契约，而不被迫统一输出格式。

## 兼容 facade 策略

`utils.py` 可以经历两个阶段：

### 阶段 A：内部提取

先在 `utils.py` 内部提取 helper，不改变任何 import。

### 阶段 B：模块迁移

迁移实现到新模块后，`utils.py` 仅保留：

```python
from file_io import find_file_path, read_file_with_appropriate_method, write_result_file
from matching import extract_p_number, match_orders_by_p_number, process_excel_files
from sales_report import (
    add_sales_report_period,
    filter_unmarked_and_generate_report,
    get_year_month,
    parse_date,
    process_sales_report_workflow,
)
```

这允许现有入口和第三方脚本继续使用旧导入路径。

## 测试策略

先写行为锁定测试，再重构实现。

重点测试矩阵：

| 区域 | 必测行为 |
|------|----------|
| matching | exact 优先、P-number、hyphen、fallback 行顺序 |
| matching | 正单、退单、零金额 |
| matching | `process_excel_files()` 仍刷新 `销售报表账期` |
| file_io | CSV 编码 fallback、分隔符 fallback、注释行跳过 |
| file_io | `.xlsx` / `.xls` 读取路径、订单号字符串保护 |
| sales_report | 全退、已取消、未标注筛选、日期窗口 |
| workflow_service | 统计、错误归一化、持久化协调 |
| adapters | CLI/API 输出 shape 不变 |

## 不确定点和后续决策

以下问题在本 change 中只记录，不改变：

- 销售报表日期窗口目前是目标月份前后 1 年；注释中曾出现“往前一年”的说法。若要更改语义，应单独开 change。
- `--match-only` 当前是否应该继续刷新 `销售报表账期`，由 `process_excel_files()` 的现有副作用决定。若要让 match-only 真正只匹配手续费，应单独开 change。
- API 是否应迁移到 app factory、是否清理 import-time 目录创建，不包含在本 change。
- 匹配性能优化不包含在本 change，避免改变候选优先级。
