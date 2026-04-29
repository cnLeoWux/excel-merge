---
name: excel-merge-cli
description: Run the excel-merge CLI to match order files with payment/refund files; optionally trigger the sales-report workflow. Optimized for Feishu workflow - accepts two uploaded files and sends the processed result back to the chat.
license: MIT
metadata:
  author: excel-merge
  version: "3.0"
---

# Excel Merge CLI Skill (Feishu Optimized)

Use this skill when the user wants to merge Excel/CSV files in Feishu:

1. **Two files uploaded** → Match order file with payment/refund file, fill in `支付手续费` column
2. **With --month flag** → Also run sales report workflow, marking `销售报表账期` column
3. **Send result back** → Upload the processed file to the Feishu chat

> **In-place contract**: The CLI writes results back to the original order file. For Feishu workflow, we create a temporary copy to preserve the original, then send the processed copy to chat.

---

## When to Use

Trigger this skill when:
- User uploads 2 Excel/CSV files to Feishu chat
- User says "合并这两个文件" / "匹配订单和支付"
- User mentions "销售报表" / "标注账期"
- Files are in `ExcelForHandel/` folder

**Typical Feishu workflow:**
1. User uploads `order.xlsx` (订单文件)
2. User uploads `payment.xlsx` (支付/退款文件)
3. You download both files
4. Run `excel-merge-cli` on them
5. Send the processed `order.xlsx` back to chat

---

## Feishu Workflow Steps

### Step 1: Receive Files
When user uploads files to Feishu:
- Files are automatically saved to `/tmp/openclaw/` with unique names
- File info includes `file_key` for downloading

### Step 2: Download Files
Use `feishu_im_bot_image` to download uploaded files:
```python
# For each uploaded file
feishu_im_bot_image(
    message_id=message_id,
    file_key=file_key,
    type="file"  # or "image" for screenshots
)
```

### Step 3: Identify File Types

**Order file** (订单文件) typically contains:
- `订单号` column
- `外部订单号` column
- `订单金额` column
- `商品名称` column (for P-number matching)

**Payment file** (支付/退款文件) typically contains:
- `商户订单号` or `商户`+`订单` columns
- `支出金额（-元）` or `收入金额（+元）` column
- Business type column (`收费`/`服务费`/`退费`/`退款`)

> **Auto-detection**: If unsure which is which, check column names. Order file has `订单号`, payment file has `商户订单号`.

### Step 4: Run CLI

```bash
# Basic match
python cli.py /path/to/order.xlsx /path/to/payment.xlsx --json --quiet

# With sales report workflow
python cli.py /path/to/order.xlsx /path/to/payment.xlsx --month 202602 --json --quiet
```

### Step 5: Send Result Back

Use `feishu_im_user_message` to send the processed file:
```python
feishu_im_user_message(
    action="send",
    msg_type="file",
    content=json.dumps({"file_key": uploaded_file_key}),
    receive_id_type="chat_id",
    receive_id=chat_id
)
```

---

## Argument Reference

| Argument | Required | Default | Purpose |
|---|---|---|---|
| `order_file` | yes | — | Order data file path |
| `payment_file` | yes | — | Payment/refund file path |
| `--month YYYYMM` | no | `None` | Trigger sales report workflow |
| `--json` | no | `False` | Emit JSON output (recommended for automation) |
| `--quiet` | no | `False` | Suppress progress logs |

---

## Canonical Invocations

### Basic Match (Feishu workflow)
```bash
# Create temp copy to preserve original
cp /tmp/openclaw/order.xlsx /tmp/openclaw/order_processed.xlsx

# Run CLI
python cli.py /tmp/openclaw/order_processed.xlsx /tmp/openclaw/payment.xlsx --json --quiet

# Result is in order_processed.xlsx, ready to send back
```

### Sales Report Workflow
```bash
cp /tmp/openclaw/order.xlsx /tmp/openclaw/order_processed.xlsx
python cli.py /tmp/openclaw/order_processed.xlsx /tmp/openclaw/payment.xlsx --month 202602 --json --quiet
```

---

## JSON Output Shape

**Success:**
```json
{
  "ok": true,
  "data": {
    "output_file": "/tmp/openclaw/order_processed.xlsx",
    "statistics": {
      "total_rows": 100,
      "matched_rows": 85,
      "match_rate": "85.00%"
    }
  },
  "error": null
}
```

**Failure:**
```json
{
  "ok": false,
  "data": null,
  "error": {
    "code": "file_not_found",
    "message": "File 'x.xlsx' does not exist."
  }
}
```

---

## Feishu-Specific Implementation

### Complete Workflow Example

```python
import json
from pathlib import Path

# 1. Files are uploaded to Feishu, get their paths
order_path = "/tmp/openclaw/order_xxx.xlsx"
payment_path = "/tmp/openclaw/payment_xxx.xlsx"

# 2. Create processed copy
import shutil
processed_path = "/tmp/openclaw/order_processed.xlsx"
shutil.copy(order_path, processed_path)

# 3. Run CLI
import subprocess
result = subprocess.run(
    [
        "python", "cli.py",
        processed_path,
        payment_path,
        "--json", "--quiet"
    ],
    capture_output=True,
    text=True,
    cwd="/path/to/excel-merge"
)

# 4. Parse result
output = json.loads(result.stdout)
if output["ok"]:
    stats = output["data"]["statistics"]
    print(f"✅ 匹配成功: {stats['matched_rows']}/{stats['total_rows']} ({stats['match_rate']})")
    
    # 5. Upload processed file back to Feishu
    # (Use feishu_drive_file upload to get file_key, then send message)
else:
    print(f"❌ 错误: {output['error']['message']}")
```

### Sending File Back to Feishu Chat

```python
# Upload file to Feishu to get file_key
feishu_drive_file(
    action="upload",
    file_path="/tmp/openclaw/order_processed.xlsx",
    folder_token="your_folder_token"  # Optional
)

# Send file message
feishu_im_user_message(
    action="send",
    msg_type="file",
    content=json.dumps({"file_key": file_key}),
    receive_id_type="chat_id",
    receive_id="oc_xxx"  # Group chat ID
)
```

---

## Exit Codes & Error Handling

| Code | Meaning | Feishu Response |
|---|---|---|
| 0 | Success | Send processed file + stats |
| 1 | General error | Reply with error message |
| 2 | Usage error | Check command syntax |
| 3 | File not found | Ask user to re-upload |
| 4 | Processing error | Check file format/columns |

---

## Common Pitfalls (Feishu Context)

- **File locked**: If user has order file open in Excel, writing will fail. Ask them to close it first.
- **Wrong file order**: First arg must be order file, second is payment file. Auto-detect by column names if unsure.
- **CSV encoding**: Auto-detected (gbk → utf-8 → gb2312 → latin-1 → utf-8-sig)
- **20-char truncation**: Order numbers are matched by first 20 chars only
- **P-number case-sensitive**: Must be uppercase `P\d+`, lowercase won't match

---

## Quick Reference

```bash
# Help
python cli.py --help

# Basic match (Feishu optimized)
cp order.xlsx order_processed.xlsx
python cli.py order_processed.xlsx payment.xlsx --json --quiet

# Sales report
cp order.xlsx order_processed.xlsx
python cli.py order_processed.xlsx payment.xlsx --month 202602 --json --quiet
```

---

## Feishu Message Templates

**Success response:**
```
✅ Excel 合并完成！

📊 匹配统计：
   • 总订单数：{total_rows}
   • 成功匹配：{matched_rows}
   • 匹配率：{match_rate}

📎 已处理文件已上传
```

**With sales report:**
```
✅ Excel 合并 + 销售报表标注完成！

📊 匹配统计：
   • 总订单数：{total_rows}
   • 成功匹配：{matched_rows}
   • 匹配率：{match_rate}

📝 销售报表：
   • 目标月份：{month}
   • 已标注账期信息

📎 已处理文件已上传
```

**Error response:**
```
❌ 处理失败：{error_message}

请检查：
   • 文件格式是否正确 (.xlsx/.xls/.csv)
   • 订单文件是否包含"订单号"列
   • 支付文件是否包含"商户订单号"列
   • 文件是否被其他程序占用
```
