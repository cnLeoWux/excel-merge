# Excel Merge Tool - 使用文档

## 📋 项目简介

Excel Merge Tool 是一个用于**订单数据与支付流水自动匹配**的工具，根据订单号、P-number 等规则匹配支付手续费。同时支持销售报表账期标记和月度报表生成。

---

## 🚀 四种使用方式

### 方式一：交互式命令行（原版）

```bash
python excel_merge.py
# 或安装后：
excel-merge
```

按提示选择 `ExcelForHandel/` 目录下的文件进行交互式处理。

---

### 方式二：命令行参数模式（推荐批量处理）

```bash
# 基本用法 - 直接修改原订单文件
python cli.py <订单文件> <支付流水文件>

# 指定输出文件
python cli.py orders.xlsx payments.xlsx -o result.xlsx

# 销售报表工作流
python cli.py orders.xlsx payments.xlsx --month 202602

# 销售报表 + 指定输出目录
python cli.py orders.xlsx payments.xlsx --month 202602 --output-dir ./reports

# 示例
python cli.py ./data/orders_202403.xlsx ./data/payments_202403.xlsx -o ./output/merged.xlsx
```

**支持的格式**：`.xlsx`, `.xls`, `.csv`

**CLI 参数一览**：

| 参数 | 说明 | 必填 |
|------|------|------|
| `order_file` | 订单数据文件路径 | 是 |
| `payment_file` | 支付流水文件路径 | 是 |
| `-o`, `--output` | 输出文件路径（默认覆盖原文件） | 否 |
| `--month` | 目标月份 YYYYMM，触发销售报表工作流 | 否 |
| `--output-dir` | 报表输出目录 | 否 |
| `--json` | 以 JSON 格式输出结果到 stdout | 否 |
| `--quiet` | 静默模式，仅输出错误信息 | 否 |
| `-v`, `--verbose` | 详细日志模式（-v=INFO, -vv=DEBUG） | 否 |

---

### 方式二（扩展）：AI Agent / 自动化调用模式

适用于 AI Agent（如 Claude Code、Cursor）或自动化脚本调用：

```bash
# JSON 输出 + 静默模式（推荐用于自动化）
python cli.py orders.xlsx payments.xlsx --json --quiet

# 检查处理结果和退出码
python cli.py orders.xlsx payments.xlsx --json --quiet
# 输出示例：
# {"ok": true, "data": {"output_file": "orders.xlsx", "statistics": {...}}, "error": null}
echo $?  # 0 表示成功

# 文件不存在时的错误处理
python cli.py nonexistent.xlsx payments.xlsx --json --quiet
# 输出：{"ok": false, "data": null, "error": {"code": "file_not_found", "message": "..."}}
echo $?  # 3 表示文件不存在
```

**退出码说明**：

| 退出码 | 含义 | 使用场景 |
|--------|------|----------|
| 0 | 成功 | 处理完成，结果已输出 |
| 1 | 通用错误 | 未预期的异常 |
| 2 | 用法错误 | 参数无效或缺失 |
| 3 | 文件未找到 | 输入文件不存在 |
| 4 | 处理错误 | 匹配或写入过程中出错 |

**非交互式模式**（适用于 excel_merge.py）：

```bash
# 使用 excel_merge.py 的非交互式模式
python excel_merge.py --non-interactive \
  --order-file orders.xlsx \
  --payment-file payments.xlsx \
  --json --quiet

# 自动检测非 TTY 环境（如管道输入）
echo "" | python excel_merge.py --order-file orders.xlsx --payment-file payments.xlsx
```

---

### 方式三：HTTP API 服务（支持钉钉集成）

#### 1. 启动服务

```bash
# 安装依赖
pip install -r requirements.txt

# 启动 API 服务器
python excel_merge_api.py
```

服务将在 `http://localhost:5000` 启动。

#### 2. API 端点

| 方法 | 端点 | 说明 |
|------|------|------|
| GET | `/` | Web 测试页面 |
| GET | `/health` | 健康检查 |
| POST | `/merge` | 上传文件，直接返回处理后的文件 |
| POST | `/merge/json` | 上传文件，返回 JSON（含下载链接） |
| GET | `/download/<filename>` | 下载结果文件 |

#### 3. 调用示例

**cURL 直接下载结果：**
```bash
curl -X POST http://localhost:5000/merge \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.xlsx" \
  --output merged_result.xlsx
```

**cURL JSON 模式（获取下载链接）：**
```bash
curl -X POST http://localhost:5000/merge/json \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.xlsx"
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
    "payment": "payments.xlsx",
    "result": "merged_result_20260330_143022_a1b2c3d4.xlsx"
  }
}
```

**Python requests 示例：**
```python
import requests

url = "http://localhost:5000/merge"
files = {
    'order_file': open('orders.xlsx', 'rb'),
    'payment_file': open('payments.xlsx', 'rb')
}

response = requests.post(url, files=files)
with open('result.xlsx', 'wb') as f:
    f.write(response.content)
```

---

### 方式四：控制台命令（安装后）

```bash
# 安装为可编辑包
pip install -e .

# 交互式模式
excel-merge

# CLI 模式
excel-merge-cli orders.xlsx payments.xlsx -o result.xlsx
excel-merge-cli orders.xlsx payments.xlsx --month 202602
```

控制台命令通过 `setup.py` 的 `entry_points` 注册：
- `excel-merge` → `excel_merge:main`（交互式）
- `excel-merge-cli` → `cli:main_cli`（命令行）

---

## 📊 销售报表工作流

通过 `--month` 参数触发完整的销售报表生成流程：

```bash
python cli.py orders.xlsx payments.xlsx --month 202602 --output-dir ./reports
```

### 处理流程

**第一阶段：匹配与标记**
1. 匹配支付手续费（与基本模式相同）
2. 在"销售报表账期"列标记特殊订单：
   - **全退**：同一订单号出现多次，金额合计为0
   - **已取消**：订单状态含"取消"且金额为0

**第二阶段：筛选与生成报表**
1. 过滤掉已标记的行（全退、已取消）
2. 筛选"出行日期"在目标月份前1年范围内的数据
   - 例如 `--month 202602` → 筛选 2025-02-01 至 2026-02-28 的出行日期
3. 在原数据中标记为"销售报表202602"
4. 生成新文件 `report_202602.xlsx`，包含筛选出的数据

---

## 🔧 钉钉集成方案

### 方案 A：通过 Jarvis 中转（推荐）

1. **用户**在钉钉群上传两个 Excel 文件
2. **Jarvis** 接收文件并保存到临时目录
3. **Jarvis** 调用本地 API：`curl -F "order_file=@..." -F "payment_file=@..." http://localhost:5000/merge`
4. **Jarvis** 将结果文件回传到钉钉群

### 方案 B：独立部署

将 API 服务部署到内网服务器，钉钉机器人直接调用 HTTP 接口。

---

## 📁 文件说明

| 文件 | 用途 |
|------|------|
| `excel_merge.py` | 交互式命令行版本 |
| `cli.py` | 命令行参数版本 |
| `excel_merge_api.py` | HTTP API 服务 |
| `utils.py` | 核心处理逻辑（~930行） |
| `setup.py` | 包配置，注册控制台命令 |
| `requirements.txt` | Python 依赖 |
| `USAGE.md` | 本文档 |
| `README.md` | 英文项目概览 |
| `uploads/` | API 上传文件暂存目录（自动创建） |
| `results/` | API 处理结果目录（自动创建） |

---

## ⚙️ 匹配逻辑说明

1. **正单**（订单金额 > 0）：匹配「收费」或「服务费」类型记录，手续费 = 支出金额（负值）
2. **退单**（订单金额 < 0）：匹配「退费」或「退款」类型记录，手续费 = 收入金额（正值）
3. **零金额**：手续费设为 0

**匹配优先级**：
1. 订单号前20位精确匹配
2. P-number 匹配（外部订单号 vs 商品名称中的 `P\d+` 模式）
3. 连字符分隔符匹配（外部订单号 vs 商品名称最后 `-` 之后的部分）

---

## 📝 注意事项

- 确保输入文件包含必要的列：订单号、外部订单号、订单金额、业务类型等
- CSV 文件支持多种编码（自动检测 gbk、utf-8、gb2312、latin-1、utf-8-sig）
- 不指定 `-o` 时默认修改原文件，建议先备份
- API 服务默认保存上传和处理记录，可定期清理 `uploads/` 和 `results/` 目录
- 控制台命令需先运行 `pip install -e .` 注册
