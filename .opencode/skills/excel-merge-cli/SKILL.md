---
name: excel-merge-cli
description: Run the excel-merge CLI to match order files with payment/refund files; optionally trigger the sales-report workflow. Optimized for Feishu workflow - accepts two uploaded files and sends the processed result back to the chat.
license: MIT
metadata:
  author: excel-merge
  version: "3.2"
---

# Excel Merge CLI Skill (Feishu Optimized)

Use this skill when the user wants to merge Excel/CSV files in Feishu:

1. **Two files uploaded** → Match order file with payment/refund file, fill in `支付手续费` column
2. **Auto-detect month from filename** → Extract YYYYMM pattern (e.g., 202603 from filename)
3. **Confirm or ask user** → Use a reliably detected month; if no reliable month is found, ask the user before processing
4. **With target_month argument** → Also run sales report workflow, marking `销售报表账期` column
5. **Send result back** → Upload the processed file to the Feishu chat

> **In-place contract**: The CLI writes results back to the original order file. For Feishu workflow, we create a temporary copy to preserve the original, then send the processed copy to chat.

---

## When to Use

Trigger this skill when:
- User uploads 2 Excel/CSV files to Feishu chat (even without saying anything, or just saying "这两个", "弄一下")
- User uses colloquial terms for merging: "对个账", "合并一下", "匹配订单和支付", "算一下手续费"
- User mentions report generation: "做个报表", "X月的账期", "销售报表"
- Files are in `ExcelForHandel/` folder or recently uploaded in the chat context.

---

## ⚠️ Critical Implementation Notes

### 1. 获取正确的 message_id（文件消息，而非请求消息）

**错误做法**：直接使用合并请求的 message_id 调用 feishu_im_bot_image
**正确做法**：先调用 `feishu_im_user_get_messages` 获取最近消息，识别文件消息的 message_id

```python
# Step 1: 获取最近消息列表
messages = feishu_im_user_get_messages(
    chat_id="oc_xxx",
    page_size=10,
    sort_rule="create_time_desc"
)

# Step 2: 找到文件消息（msg_type == "file"）
for msg in messages["messages"]:
    if msg["msg_type"] == "file":
        # 解析 content 获取 file_key 和文件名
        # content 格式: <file key="file_v3_xxx" name="文件名.xlsx"/>
        message_id = msg["message_id"]
        file_key = parse_file_key(msg["content"])
```

### 2. CSV 文件必须添加 .csv 扩展名

**问题**：飞书下载的文件可能没有扩展名，导致 CLI 把 CSV 当 Excel 处理
**解决**：下载后检查并添加正确的扩展名

```python
# 下载文件
result = feishu_im_bot_image(
    message_id=message_id,
    file_key=file_key,
    type="file"
)

saved_path = result["saved_path"]

# 检查文件类型并添加扩展名
import subprocess
file_type = subprocess.run(["file", saved_path], capture_output=True, text=True).stdout

if "CSV" in file_type or "text" in file_type and not saved_path.endswith(".csv"):
    new_path = saved_path + ".csv"
    os.rename(saved_path, new_path)
    saved_path = new_path
elif "Excel" in file_type and not saved_path.endswith((".xlsx", ".xls")):
    new_path = saved_path + ".xlsx"
    os.rename(saved_path, new_path)
    saved_path = new_path
```


### 3. Default workflow requires target_month
The default and preferred workflow is the full workflow: payment fee matching + sales report period marking + date-window marking. Therefore `target_month` is required for normal Skill execution.

- First infer the month from filenames or conversation context.
- If inference is reliable, use that month directly or briefly confirm when ambiguity exists.
- If inference is not reliable, ask the user for the month before invoking the CLI.
- Do not silently run `--match-only` just because the month is missing.
- Only run `--match-only` when the user explicitly asks for matching-only/no sales report/no period marking.

### 4. Native CSV Robustness
The application natively handles CSV edge cases (like `="1234"` prefixes, long integer float coercion, and bad lines). However, **you must ensure the file has a `.csv` extension** so the CSV engine is triggered.

### 5. CLI 执行路径与解释器

**问题**：
- exec preflight 阻止 `cd && python` 组合命令
- 系统只有 `python3` 没有 `python`

**解决**：使用 `/usr/bin/python3` 绝对路径，单独执行

```bash
# 正确方式
/usr/bin/python3 /path/to/excel-merge/cli.py order.xlsx payment.csv 202603 --match-only --json --quiet

# 错误方式（会被 exec preflight 阻止）
cd /path/to/excel-merge && python cli.py ...
```

### 6. 文件发送到群组

**方案 A**：使用 message 工具的 buffer 参数发送 base64 文件
```python
import base64
file_content = base64.b64encode(open(processed_file, "rb").read()).decode()

message(
    action="send",
    channel="feishu",
    target="oc_xxx",  # 群组 ID
    filename="订单数据_已合并.xlsx",
    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer=f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{file_content}"
)
```

**方案 B**：上传到云空间后分享
```python
# 上传到云空间
upload_result = feishu_drive_file(
    action="upload",
    file_path=processed_file,
    folder_token=""  # 我的空间根目录
)

# 发送文件卡片到群组（需要 file_token）
# 注意：目前 feishu_drive_file upload 返回结果可能不完整
```

---

## Feishu Workflow Steps (完整流程)

### Step 0: 智能意图识别与默认全量执行 (Smart Intent & Default Full Workflow)
为了保证最大的处理效益，**本 SKILL 默认且最优先执行包含匹配和标注账期的“完整流程（第一和第二阶段）”**。在开始执行前，按以下逻辑判断和沟通：

1. **默认全量意图（看破不说破，直接推导完整流程）**：
   - 无论用户上传两个文件时说什么（“合并”、“对账”、“这两个文件”、“弄一下”甚至是空白），请**一律默认**用户希望执行【完整工作流】（即既做匹配，又做销售报表账期标注）。
2. **主动索要或推断月份（触发完整流程的前提）**：
   - 既然默认执行完整流程，那么 `target_month` 参数是必须的。
   - **推断**：先自动观察文件名或当前聊天上下文是否包含自然语言日期（如“3月份的数据”、“上个月”、“2026年3月”），如果有，请自动转为 `YYYYMM`（如 202603）。
   - **索要**：如果无法推断，在开始前**务必主动询问**：“已收到订单和支付文件，为了帮您执行完整的合并与账期标注流程，请问需要处理哪个月份的数据？（例如：202603，或直接回复‘上个月’）”
3. **检查文件齐备度**：
   - 流程严格需要 **两个** 文件。如果用户只发了一个，主动提示：“我还缺少对应的文件（需要同时提供订单数据和支付流水）才能执行完整流程，请您补充上传。”
4. **不盲目降级**：
   - 除非用户在对话中明确且强烈地表示“**不要**做销售报表/账期”、“我**只要**单纯合并金额”，否则**不要**降级去只做 `--match-only`。默认情况下必须想办法获取月份去执行完整命令。

### Step 1: 获取文件消息列表

```python
messages = feishu_im_user_get_messages(
    chat_id=chat_id,
    page_size=10,
    sort_rule="create_time_desc"
)

file_messages = []
for msg in messages["messages"]:
    if msg["msg_type"] == "file":
        # 解析 content: <file key="xxx" name="文件名"/>
        import re
        match = re.search(r'key="([^"]+)" name="([^"]+)"', msg["content"])
        if match:
            file_messages.append({
                "message_id": msg["message_id"],
                "file_key": match.group(1),
                "file_name": match.group(2)
            })
```

### Step 2: 下载文件并添加扩展名

```python
order_path = None
payment_path = None

for fm in file_messages:
    result = feishu_im_bot_image(
        message_id=fm["message_id"],
        file_key=fm["file_key"],
        type="file"
    )
    
    saved_path = result["saved_path"]
    
    # 添加扩展名（如果缺失）
    if fm["file_name"].endswith(".csv") and not saved_path.endswith(".csv"):
        saved_path = saved_path + ".csv"
        os.rename(result["saved_path"], saved_path)
    elif fm["file_name"].endswith(".xlsx") and not saved_path.endswith(".xlsx"):
        saved_path = saved_path + ".xlsx"
        os.rename(result["saved_path"], saved_path)
    
    # 识别文件类型（订单 vs 支付）
    if "订单" in fm["file_name"]:
        order_path = saved_path
    elif "账务" in fm["file_name"] or "明细" in fm["file_name"]:
        payment_path = saved_path
```

### Step 3: 自动识别并确认月份 (为了执行全量流程)

```python
# 从文件名自动提取月份
import re

month_pattern = re.search(r'20[0-9]{4}', payment_file_name)
if month_pattern:
    target_month = month_pattern.group()
else:
    # 结合 Step 0 向用户主动询问获取月份
    target_month = ask_user_for_month()
```

### Step 4: 执行 CLI

```bash
CLI_PATH="$(pwd)/cli.py"  # Ensure execution from the project root

/usr/bin/python3 $CLI_PATH $order_path $payment_path $target_month --json --quiet  # 默认执行完整流程
```

### Step 5: 发送结果到群组

```python
# 方案 A：使用 message 工具的 buffer 参数
import base64

file_content = base64.b64encode(open(order_path, "rb").read()).decode()

message(
    action="send",
    channel="feishu",
    target=chat_id,
    filename="订单数据_已合并.xlsx",
    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer=f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{file_content}"
)
```

---

## Argument Reference

| Argument | Required | Default | Purpose |
|---|---|---|---|
| `order_file` | yes | — | Order data file path |
| `payment_file` | yes | — | Payment/refund file path |
| `target_month` | positional optional in argparse, required for default Skill workflow | — | 位置参数月份 (YYYYMM)。默认完整流程必须提供；缺失时先推断，无法推断则询问用户 |
| `--match-only` | no | — | 显式降级：仅执行匹配（需要用户明确要求，且当前 CLI 仍需要 target_month） |
| `--mark-only` | no | — | 仅执行标注（需要 target_month） |
| `--json` | no | `False` | Emit JSON output |
| `--quiet` | no | `False` | Suppress progress logs |
| `-v` / `-vv` | no | — | 详细日志模式 |

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

## Exit Codes & Error Handling

| Code | Meaning | Action |
|---|---|---|
| 0 | Success | 发送文件 + 统计信息到群组 |
| 1 | General error | 返回错误信息 |
| 2 | Usage error | 检查命令参数 |
| 3 | File not found | 重新下载文件 |
| 4 | Processing error | 检查文件格式/扩展名 |

---

## Common Pitfalls

| 问题 | 原因 | 解决方案 |
|---|---|---|
| `BadZipFile: File is not a zip file` | CSV 文件无 `.csv` 扩展名，被当作 Excel 处理 | 下载后添加 `.csv` 扩展名 |
| `command not found: python` | 系统只有 `python3` | 使用 `/usr/bin/python3` |
| `exec preflight: complex interpreter invocation` | `cd && python` 组合命令 | 单独执行，不组合 |
| `Bot is NOT the owner of the resource` | 使用用户的 file_key 发送文件 | 上传新文件获取新的 file_key |
| `请输入目标月份` | 未提供 target_month 参数 | 先从文件名/上下文推断；无法推断时询问用户并提供 `202603` 格式的月份参数 |

---

## Quick Reference

```bash
# CLI 路径
CLI_PATH="$(pwd)/cli.py"  # Ensure execution from the project root

# 默认执行完整工作流（匹配 + 标注 + 日期筛选）
/usr/bin/python3 $CLI_PATH order.xlsx payment.csv 202603 --json --quiet

# 仅在用户明确要求“只匹配/不要账期/不要销售报表”时执行匹配-only
/usr/bin/python3 $CLI_PATH order.xlsx payment.csv 202603 --match-only --json --quiet
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

📎 已处理文件已发送到群组
```

**Error response:**
```
❌ 处理失败：{error_message}

请检查：
   • 文件格式是否正确 (.xlsx/.xls/.csv)
   • 订单文件是否包含"订单号"列
   • 支付文件是否包含"商户订单号"列
```
