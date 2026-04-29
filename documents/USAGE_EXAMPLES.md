# Excel Merge Tool - Usage Examples

> **⚠ 破坏性变更（迁移说明）**：CLI 已移除 `-o`/`--output` 与 `--output-dir` 参数；销售报表工作流不再产出独立的 `report_YYYYMM.xlsx` 文件。所有合并与销售报表标记结果**一律就地写回订单文件**。如果需要"另存为"效果，请先复制订单文件再调用 CLI（例如 `cp order.xlsx order_copy.xlsx && python cli.py order_copy.xlsx payment.xlsx`）。HTTP API（`excel_merge_api.py`）的对外契约不受此影响，仍可下载月报文件。

## Installation

```bash
# 安装依赖
pip install -r requirements.txt

# 可选：注册控制台命令
pip install -e .
```

安装后可使用 `excel-merge`（交互式）和 `excel-merge-cli`（CLI）命令。

---

## Interactive Mode (交互式)

```bash
python excel_merge.py
# 或安装后：
excel-merge
```

程序会列出 `ExcelForHandel/` 目录下的所有文件，按编号选择订单文件和支付文件：

```
Excel Merge Tool
Available files in ExcelForHandel directory:
1. orders_202602.xlsx
2. payments_202602.csv
3. refunds_202602.xlsx

Select the first Excel file (order data) by number: 1
Select the second Excel file (payment/refund data) by number: 2
```

处理完成后直接修改原订单文件。

---

## CLI Mode (命令行)

> CLI 只暴露一种产出渠道：**就地修改订单文件**。不会生成任何独立的结果文件或月度报表文件。如需备份，请在调用前自行复制订单文件。

### 基本用法

```bash
# 就地匹配支付手续费并写回 order.xlsx
python cli.py order.xlsx payment.xlsx

# 或安装后：
excel-merge-cli order.xlsx payment.xlsx
```

### 销售报表工作流

```bash
# 触发完整工作流：匹配 + 标注 + 在内存中筛选 + 就地回填 销售报表账期 列
python cli.py order.xlsx payment.xlsx --month 202602
```

`--month` 触发完整销售报表工作流：
1. 匹配支付手续费
2. 标记"全退"和"已取消"订单
3. 在内存中筛选出行日期在目标月份前后 1 年范围内的未标记数据
4. 将 `销售报表YYYYMM` 回填到这些行的 `销售报表账期` 列，并就地写回订单文件
5. **不**生成任何 `report_YYYYMM.xlsx` 文件

### CLI 参数一览

| 参数 | 类型 | 默认值 | 说明 | 必填 |
|------|------|--------|------|------|
| `order_file` | str | *(必填)* | 订单数据文件路径（.xlsx, .xls, .csv） | 是 |
| `payment_file` | str | *(必填)* | 支付流水文件路径（.xlsx, .xls, .csv） | 是 |
| `--month` | str | `None` | 目标月份 `YYYYMM` 格式（如 `202602`），触发销售报表工作流 | 否 |
| `--json` | flag | `False` | 以 JSON 信封格式输出结果到 stdout | 否 |
| `--quiet` | flag | `False` | 静默模式，仅输出警告和错误到 stderr | 否 |
| `-v`, `--verbose` | count | `0` | 详细日志模式（-v=INFO, -vv=DEBUG） | 否 |

> 已移除：`-o`/`--output`、`--output-dir`。传入这些参数会被 argparse 拒绝并以退出码 2 退出。

---

## AI Agent / Automation Mode (AI Agent / 自动化模式)

适用于 AI Agent（如 Claude Code、Cursor）或自动化脚本调用，提供结构化输出和语义化退出码。

### JSON Output Mode

```bash
# JSON 输出 + 静默模式（推荐用于自动化）
python cli.py order.xlsx payment.xlsx --json --quiet

# 输出示例：
{
  "ok": true,
  "data": {
    "output_file": "order.xlsx",
    "statistics": {
      "total_rows": 100,
      "matched_rows": 85,
      "match_rate": "85.00%"
    }
  },
  "error": null
}
```

无论是否传入 `--month`，`data` 的形状始终一致，仅含 `output_file`（等于订单文件路径）与 `statistics`，**不**包含 `report_file` / `report_rows` / `warnings` 字段。

### Error Handling

```bash
# 文件不存在时的错误输出
python cli.py nonexistent.xlsx payment.xlsx --json --quiet

# 输出示例：
{
  "ok": false,
  "data": null,
  "error": {
    "code": "file_not_found",
    "message": "File 'nonexistent.xlsx' does not exist."
  }
}

# 检查退出码
echo $?  # 输出: 3
```

`error.code` 的可能值: `file_not_found`、`processing_error`、`unknown_error`。

### Exit Codes

| 退出码 | 常量 | 含义 | 使用场景 |
|--------|------|------|----------|
| 0 | `EXIT_SUCCESS` | 成功 | 处理完成，结果已输出 |
| 1 | `EXIT_GENERAL_ERROR` | 通用错误 | 未预期的异常 |
| 2 | `EXIT_USAGE_ERROR` | 用法错误 | 参数无效或缺失 |
| 3 | `EXIT_FILE_NOT_FOUND` | 文件未找到 | 输入文件不存在 |
| 4 | `EXIT_PROCESSING_ERROR` | 处理错误 | 匹配或写入过程中出错 |

### stdout/stderr 分离规则

- **stdout**: 仅输出 JSON 结果（`--json` 模式）或结果文件路径（文本模式）
- **stderr**: 所有日志、进度信息、警告和错误

使用 `--json --quiet` 时，stdout 中仅包含 JSON 信封，所有其他输出均发送到 stderr，便于程序化解析。

### Non-Interactive Mode (excel_merge.py)

```bash
# 使用 excel_merge.py 的非交互式模式
python excel_merge.py --non-interactive \
  --order-file order.xlsx \
  --payment-file payment.xlsx \
  --json --quiet

# 自动检测非 TTY 环境（如管道输入）
echo "" | python excel_merge.py \
  --order-file order.xlsx \
  --payment-file payment.xlsx \
  --json

# 缺少必要参数时的错误
python excel_merge.py --non-interactive
# 输出: Error: Non-interactive mode requires --order-file and --payment-file arguments.
# 退出码: 2
```

### Python Script Integration

```python
import subprocess
import json

# Run the CLI with JSON output
result = subprocess.run(
    ["python", "cli.py", "order.xlsx", "payment.xlsx", "--json", "--quiet"],
    capture_output=True,
    text=True
)

# Parse the result
if result.returncode == 0:
    data = json.loads(result.stdout)
    if data["ok"]:
        print(f"Matched {data['data']['statistics']['matched_rows']} rows")
        print(f"Match rate: {data['data']['statistics']['match_rate']}")
    else:
        print(f"Error: {data['error']['message']}")
else:
    print(f"Process failed with exit code: {result.returncode}")
```

---

## Flask API Mode (HTTP 服务)

### 启动服务

```bash
python excel_merge_api.py
```

服务运行在 `http://localhost:5000`，提供 Web 测试页面。

### API Endpoints

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | Web 测试页面（含文件上传表单） |
| GET | `/health` | 健康检查 |
| POST | `/merge` | 上传文件，直接返回处理后的文件 |
| POST | `/merge/json` | 上传文件，返回 JSON（含下载链接和统计） |
| GET | `/download/<filename>` | 下载结果文件 |

### cURL: 直接下载结果

```bash
curl -X POST http://localhost:5000/merge \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.xlsx" \
  --output merged_result.xlsx
```

### cURL: JSON 模式

```bash
curl -X POST http://localhost:5000/merge/json \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.csv"
```

返回示例：
```json
{
  "success": true,
  "session_id": "a1b2c3d4",
  "download_url": "/download/merged_result_20260330_143022_a1b2c3d4.xlsx",
  "statistics": {
    "total_rows": 150,
    "matched_rows": 142,
    "match_rate": "94.7%"
  },
  "files": {
    "order": "orders.xlsx",
    "payment": "payments.csv",
    "result": "merged_result_20260330_143022_a1b2c3d4.xlsx"
  }
}
```

### cURL: 下载文件

```bash
curl http://localhost:5000/download/merged_result_20260330_143022_a1b2c3d4.xlsx \
  --output result.xlsx
```

### cURL: 健康检查

```bash
curl http://localhost:5000/health
```

```json
{
  "status": "healthy",
  "timestamp": "2026-03-30T14:30:22.123456",
  "service": "excel-merge-api"
}
```

### Python requests 示例

```python
import requests

# 方式1：直接获取文件
url = "http://localhost:5000/merge"
files = {
    'order_file': open('orders.xlsx', 'rb'),
    'payment_file': open('payments.xlsx', 'rb')
}
response = requests.post(url, files=files)
with open('result.xlsx', 'wb') as f:
    f.write(response.content)

# 方式2：JSON 模式
url_json = "http://localhost:5000/merge/json"
response = requests.post(url_json, files=files)
data = response.json()
print(f"匹配率: {data['statistics']['match_rate']}")

# 下载结果
download_url = f"http://localhost:5000{data['download_url']}"
result = requests.get(download_url)
with open('result.xlsx', 'wb') as f:
    f.write(result.content)
```

---

## Sample Data Formats

### 订单文件

| 订单号 | 外部订单号 | 订单金额 | 订单状态 | 出行日期 | 支付手续费 |
|--------|-----------|---------|---------|---------|-----------|
| 40250702110303185340xx | P2507021103060001 | 100.00 | 已确认 | 2025-07-15 | (待填充) |
| 40250701232642050749xx | P2507012326430003 | -50.00 | 已退款 | 2025-07-10 | (待填充) |
| 40250709224606388514xx | P2507092246080005 | 0.00 | 已取消 | 2025-07-20 | (待填充) |

### 支付流水文件

| 商户订单号 | 商品名称 | 业务类型 | 支出金额（-元） | 收入金额（+元） |
|-----------|---------|---------|---------------|---------------|
| 40250702110303185340yy | 吉祥旅游支付订单-P2507021103060001 | 收费 | -2.50 | |
| 40250701232642050749yy | 吉祥旅游支付订单-P2507012326430003 | 退费 | | 1.20 |

注意：
- 商户订单号列通过列名包含"商户"+"订单"自动定位
- 金额列名使用全角括号（`（` `）`）
- CSV 文件可以有 `#` 开头的注释行

---

## Matching Examples

### 场景1：精确匹配（20字符）

订单号 `40250702110303185340xx` 的前20字符 `40250702110303185340` 与商户订单号 `40250702110303185340yy` 的前20字符相同 → 匹配成功。订单金额 > 0（正单），业务类型为"收费" → 校验通过。支付手续费 = 支出金额（-元）= -2.50。

### 场景2：P-number 匹配

订单号不足20字符，回退到 P-number 匹配。从外部订单号 `P2507012326430003` 和商品名称 `吉祥旅游支付订单-P2507012326430003` 中分别提取出相同的 P-number → 匹配成功。订单金额 < 0（退单），业务类型为"退费" → 校验通过。支付手续费 = 收入金额（+元）= 1.20。

### 场景3：零金额跳过

订单金额 = 0.00 → 直接设 `支付手续费 = 0.0`，不进行匹配。

---

## Troubleshooting

### 文件找不到
```
Error: File 'order_data.xlsx' does not exist.
```
确认文件在当前工作目录或 `ExcelForHandel/` 子目录中。

### 编码错误
工具会自动按 `gbk → utf-8 → gb2312 → latin-1 → utf-8-sig` 顺序尝试。如仍失败，建议将 CSV 另存为 UTF-8 编码。

### 订单号变成科学计数法
工具已强制将订单号列转为字符串。如仍有问题，检查原始文件中该列是否被 Excel 格式化为数字。

### API 上传失败
- 检查文件扩展名：仅支持 `.xlsx`、`.xls`、`.csv`
- 检查文件大小：上限 16MB
- 确认 form field 名为 `order_file` 和 `payment_file`
