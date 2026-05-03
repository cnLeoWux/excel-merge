---
name: excel-merge-process
description: 飞书群组 Excel 订单与支付流水匹配工具 — 接收群组上传的两个 Excel 文件，自动匹配并回传结果
license: MIT
metadata:
  author: Leo Wu
  version: "2.0"
  project: excel-merge
  feishu_group: oc_ae63efc47c407a25623c0ebd73653eaf
---

使用 Excel Merge Tool CLI 处理飞书群组中用户上传的订单与支付流水文件，完成匹配、标注、筛选后自动回传结果到群组。

## 核心流程

```
群组消息（含 @bot）→ 检测到两个 Excel 文件 → 下载到本地 → 执行匹配 → 回传结果到群组
```

## 第一步：读取群组消息，识别上传的文件

当收到群组中任意用户的 @bot 消息时，按以下步骤操作：

### 1.1 获取群组最近消息

使用 `feishu_im_user_get_messages` 读取当前群组的消息：

```python
feishu_im_user_get_messages(
    chat_id="oc_ae63efc47c407a25623c0ebd73653eaf",  # 群组 ID
    page_size=20,
    sort_rule="create_time_desc"  # 最新消息在前
)
```

### 1.2 识别 Excel 文件消息

在消息列表中筛选满足以下条件的消息：
- `msg_type` 为 `file`
- `file` 类型的 content 中包含 `.xlsx` 或 `.csv` 文件名
- 同一批次（相近时间戳）内找到 **两个** Excel 文件

文件类型判断：
- 订单文件：文件名含 `订单`、`order`（不区分大小写）
- 支付文件：文件名含 `支付`、`payment`、`账务`、`流水`

如果无法通过文件名判断，以消息顺序区分：
- 第一个文件 → 订单文件
- 第二个文件 → 支付文件

### 1.3 下载文件

从文件消息的 `content` 中提取 `file_key`（格式：`file_xxx`），然后：

```python
feishu_im_user_fetch_resource(
    message_id="om_xxx",   # 消息 ID
    file_key="file_xxx",   # 从 content 中提取
    type="file"             # 文件类型
)
```

文件自动保存到 `/tmp/openclaw/excel-merge/` 目录。

### 1.4 确认文件类型

下载后检查文件内容（使用 `pandas.read_excel` 或 `csv`），根据列名判断：

| 订单文件必有列 | 支付文件必有列 |
|--------------|--------------|
| 订单号 | 商户订单号 |
| 外部订单号（可选） | 商品名称 |
| 订单金额 | 支出金额（-元）/ 收入金额（+元） |
| 订单状态 | 业务类型 |

如果两个文件都是同一类型，提示用户"检测到两个订单文件，请重新上传"。

---

## 第二步：执行匹配处理

### 2.1 确定工作模式

| 条件 | 模式 | 说明 |
|------|------|------|
| 用户未指定月份 | **交互式询问** | 通过飞书消息询问用户目标月份 |
| 用户指定月份（如"按202602月"） | 完整工作流 | `target_month=202602` |
| 用户说"只匹配"、"填充手续费" | 仅匹配 | `--match-only` |
| 用户说"标注账期" | 仅标注 | `--mark-only` |

### 2.2 交互式询问 target_month

**重要**：CLI 运行在非交互环境，如果用户未指定月份，**必须通过飞书消息询问用户**，不能依赖 CLI 的交互式提示。

当用户只上传文件但未指定月份时，发送消息询问：

```python
feishu_ask_user_question(
    questions=[{
        "header": "选择账期",
        "question": "请选择目标月份（YYYYMM），例如 202602",
        "options": [
            {"label": "2026年2月", "description": "处理2026年2月账期"},
            {"label": "2026年3月", "description": "处理2026年3月账期"},
            {"label": "2026年4月", "description": "处理2026年4月账期"},
            {"label": "2026年5月", "description": "处理2026年5月账期"}
        ],
        "multiSelect": False
    }]
)
```

或者使用普通文本消息询问：

```python
message(
    action="send",
    channel="feishu",
    target="chat:oc_ae63efc47c407a25623c0ebd73653eaf",
    message="📋 请选择目标月份（YYYYMM 格式）：\n\n例如：\n• 202602 — 2026年2月\n• 202603 — 2026年3月\n\n直接回复月份数字即可，例如：202602"
)
```

用户回复后，从消息中提取 YYYYMM 格式的月份，继续执行。

### 2.3 调用 CLI

```bash
python cli.py <订单文件> <支付文件> [target_month] [选项]
```

| 选项 | 说明 |
|------|------|
| `--match-only` | 仅匹配（填充手续费） |
| `--mark-only` | 仅标注（标记账期） |
| `--json` | JSON 格式输出 |
| `--quiet` | 静默模式（减少日志） |
| `-v` | 详细日志 |

**示例：**
```bash
# 完整工作流（指定月份）
python cli.py /tmp/order.xlsx /tmp/payment.xlsx 202602 -v

# 仅匹配
python cli.py /tmp/order.xlsx /tmp/payment.xlsx --match-only --json

# 用户说"按202603月处理"
python cli.py /tmp/order.xlsx /tmp/payment.xlsx 202603 -v
```

### 2.4 处理结果

- **覆盖原文件**：处理结果覆盖 `/tmp/openclaw/excel-merge/` 下的订单文件
- **报表文件**：完整工作流生成 `report_{target_month}.xlsx`
- **退出码**：`0=成功，2=用法错误，3=文件未找到，4=处理错误`

---

## 第三步：回传结果到群组

### 3.1 上传处理后的文件到飞书

使用 `feishu_drive_file` 上传：

```python
feishu_drive_file(
    action="upload",
    file_path="/tmp/openclaw/excel-merge/report_202602.xlsx",
    name="订单匹配报表_202602.xlsx"
)
```

### 3.2 发送结果消息到群组

使用 `message` 工具发送到群组：

```python
message(
    action="send",
    channel="feishu",
    target="chat:oc_ae63efc47c407a25623c0ebd73653eaf",
    message="✅ 匹配完成！\n\n📊 处理结果：\n- 正单匹配：X 条\n- 退单匹配：X 条\n- 未匹配：X 条\n- 全退标记：X 条\n- 已取消标记：X 条\n\n📎 报表文件已生成，请查收。"
)
```

### 3.3 发送文件消息

报表文件通过 `message` 的 `filePath` 参数发送：

```python
message(
    action="send",
    channel="feishu",
    target="chat:oc_ae63efc47c407a25623c0ebd73653eaf",
    file_path="/tmp/openclaw/excel-merge/report_202602.xlsx",
    message="📎 报表文件已生成，请查收。"
)
```

---

## 飞书交互模式

### 触发方式

群组中任意用户发送包含 **两个 Excel 文件** 的消息并 @bot，即触发处理流程。

### 关键词指令

| 用户说 | 执行的命令 |
|--------|-------|
| "匹配"、"填充手续费" | `--match-only` |
| "标注账期" | `--mark-only` |
| "按202602月" | `target_month=202602` |
| "处理"、"开始" | 完整工作流（交互式输入月份） |
| "重新处理"、"再来一次" | 清除缓存，重新执行 |

### 错误处理

| 情况 | 处理方式 |
|------|---------|
| 只上传了 1 个文件 | 回复："请上传两个文件：订单文件和支付文件" |
| 文件类型无法识别 | 回复："无法识别的文件类型，请确认文件名包含'订单'或'支付'" |
| 处理失败 | 回复："处理失败：[错误原因]，请重新上传" |
| 超时（30s 内未处理完） | 回复："处理时间较长，请在稍后查看结果" |

### 完整交互示例

```
用户: @bot [订单文件.xlsx] [支付文件.xlsx]
     "按202602月处理"

小科: 📥 已收到两个文件，正在处理...
     - 订单文件：订单文件.xlsx（324 条记录）
     - 支付文件：支付文件.xlsx（518 条记录）
     - 目标月份：202602

小科: ⚙️ 执行匹配中...

小科: ✅ 处理完成！
     📊 匹配结果：
     - 正单匹配：289 条
     - 退单匹配：35 条
     - 未匹配：12 条

     📋 标注结果：
     - 全退：3 条
     - 已取消：1 条

小科: [发送报表文件]
```

---

## 输入文件格式

### 订单文件列

| 列名 | 说明 |
|------|------|
| 订单号 | 主键，唯一标识订单 |
| 外部订单号 | P-number 等外部标识（用于 P-number 匹配） |
| 订单金额 | 正数=正单，负数=退单 |
| 订单状态 | "已确认"/"已退款"/"已取消"等 |
| 支付手续费 | 待填充（匹配后更新） |
| 销售报表账期 | 待填充（标注后更新） |

### 支付文件列

| 列名 | 说明 |
|------|------|
| 商户订单号 | 20 字符匹配键 |
| 商品名称 | 含 P-number 或连字符匹配 |
| 业务类型 | "收费"(正单) / "退费"(退单) |
| 支出金额（-元） | 负数，正单使用 |
| 收入金额（+元） | 正数，退单使用 |

---

## 三种匹配策略

| 策略 | 条件 | 匹配键 |
|------|------|--------|
| 精确匹配 | 订单号前 20 字符 = 商户订单号 | `订单号[:20]` vs `商户订单号` |
| P-number 匹配 | 外部订单号含 P-number | `外部订单号` vs `商品名称` 中的 P-number |
| 连字符匹配 | 外部订单号含连字符 | `外部订单号` 中的连字符后内容 vs `商品名称` |

---

## 注意事项

1. **业务类型校验**：正单仅匹配"收费"，退单仅匹配"退费"
2. **零金额跳过**：金额为 0 的订单不参与匹配，直接设 fee = 0
3. **日期筛选**：完整工作流按 `出行日期` 在 `target_month` 内筛选
4. **自动备份**：处理前自动备份原订单文件到 `backup/` 目录
5. **文件清理**：处理完成后删除 `/tmp/openclaw/excel-merge/` 下的临时文件
6. **群组限制**：仅处理 `oc_ae63efc47c407a25623c0ebd73653eaf` 群组的消息

---

## 环境要求

- Python 3.8+
- pandas, openpyxl, xlrd
- CLI 路径：`/Users/leowu-macmini/.openclaw/workspace-coding/excel-merge/cli.py`
