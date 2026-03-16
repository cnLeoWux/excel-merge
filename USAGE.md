# Excel Merge Tool - 使用文档

## 📋 项目简介

Excel Merge Tool 是一个用于**订单数据与支付流水自动匹配**的工具，根据订单号、P-number 等规则匹配支付手续费。

---

## 🚀 三种使用方式

### 方式一：交互式命令行（原版）

```bash
python excel_merge.py
```

按提示选择 `ExcelForHandel/` 目录下的文件进行交互式处理。

---

### 方式二：命令行参数模式（推荐批量处理）

```bash
# 基本用法 - 直接修改原订单文件
python cli.py <订单文件> <支付流水文件>

# 指定输出文件
python cli.py orders.xlsx payments.xlsx -o result.xlsx

# 示例
python cli.py ./data/orders_202403.xlsx ./data/payments_202403.xlsx -o ./output/merged.xlsx
```

**支持的格式**：`.xlsx`, `.xls`, `.csv`

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
  "download_url": "/download/merged_result_20240304_143022_a1b2c3d4.xlsx",
  "statistics": {
    "total_rows": 150,
    "matched_rows": 142,
    "match_rate": "94.7%"
  },
  "files": {
    "order": "orders.xlsx",
    "payment": "payments.xlsx",
    "result": "merged_result_20240304_143022_a1b2c3d4.xlsx"
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
| `excel_merge.py` | 交互式命令行版本（原版保留） |
| `cli.py` | 命令行参数版本 |
| `excel_merge_api.py` | HTTP API 服务 |
| `utils.py` | 核心处理逻辑 |
| `requirements.txt` | Python 依赖 |
| `USAGE.md` | 本文档 |
| `uploads/` | API 上传文件暂存目录（自动创建） |
| `results/` | API 处理结果目录（自动创建） |

---

## ⚙️ 匹配逻辑说明

1. **正单**（订单金额 > 0）：匹配「收费」类型记录，手续费 = 支出金额（负值）
2. **退单**（订单金额 < 0）：匹配「退费/退款」类型记录，手续费 = 收入金额（正值）
3. **零金额**：手续费设为 0

**匹配优先级**：
1. 订单号前20位精确匹配
2. P-number 匹配（外部订单号 vs 商品名称）
3. 连字符分隔符匹配

---

## 📝 注意事项

- 确保输入文件包含必要的列：订单号、外部订单号、订单金额、业务类型等
- CSV 文件支持多种编码（自动检测 UTF-8、GBK、GB2312）
- API 服务默认保存上传和处理记录，可定期清理 `uploads/` 和 `results/` 目录
