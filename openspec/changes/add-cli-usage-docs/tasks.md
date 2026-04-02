# Implementation Tasks

## Task 1: 在 AGENTS.md 中新增 CLI USAGE REFERENCE 章节

**描述**: 在 AGENTS.md 的 `## BUILD / LINT / TEST COMMANDS` 章节之后、`## WHERE TO LOOK` 章节之前插入新章节 `## CLI USAGE REFERENCE`。

**详细内容**:

1. **参数表格**（覆盖 8 个参数）:
   - 表头: Parameter | Type | Default | Description
   - 位置参数: `order_file` (str, required), `payment_file` (str, required)
   - 可选参数: `-o/--output`, `--month`, `--output-dir`, `--json`, `--quiet`, `-v/--verbose`
   - 每个参数包含类型、默认值、说明

2. **基本匹配工作流示例**:
   ```bash
   # Modify original file in-place
   python cli.py order.xlsx payment.xlsx
   
   # Specify output file
   python cli.py order.xlsx payment.xlsx -o result.xlsx
   ```

3. **销售报表工作流示例**:
   ```bash
   # Trigger sales report workflow
   python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports
   ```
   - 说明两阶段处理流程（匹配+标记 → 筛选+生成报表）
   - 说明输出文件命名规则 `report_YYYYMM.xlsx`

4. **JSON 输出格式**:
   - 成功信封示例（包含 ok, data, error 字段）
   - 失败信封示例（包含 ok, data, error 字段）
   - 说明 error.code 可能值: `file_not_found`, `processing_error`, `unknown_error`

5. **退出码表**:
   | Exit Code | Constant | Meaning | Scenario |
   |-----------|----------|---------|----------|
   | 0 | EXIT_SUCCESS | Success | Processing completed |
   | 1 | EXIT_GENERAL_ERROR | General Error | Unexpected exception |
   | 2 | EXIT_USAGE_ERROR | Usage Error | Invalid arguments |
   | 3 | EXIT_FILE_NOT_FOUND | File Not Found | Input file missing |
   | 4 | EXIT_PROCESSING_ERROR | Processing Error | Matching/writing error |

6. **Agent 推荐用法**:
   - 推荐使用 `--json --quiet` 组合
   - 说明 stdout 仅输出 JSON
   - 说明 stderr 输出日志
   - 提供 Python subprocess 示例

7. **stdout/stderr 分离规则**:
   - stdout: JSON result (`--json` mode) or file path (text mode)
   - stderr: logs, progress, warnings, errors
   - 说明如何使用 `capture_output=True` 分别获取

8. **常见错误场景**:
   - File not found (exit code 3)
   - Processing error (exit code 4)
   - 每个场景包含错误示例和解决建议

**验证步骤**:
1. 检查新章节插入位置正确（在 BUILD 之后、WHERE TO LOOK 之前）
2. 对比 cli.py L64-115 中的 argparse 定义，确认参数列表完整
3. 对比 cli.py L18-23 中的退出码常量，确认数值一致
4. 对比 cli.py L26-59 中的 `output_result()` 函数，确认 JSON 格式一致
5. 手动运行 `python cli.py --help` 确认参数说明与实际一致
6. 手动运行 `python cli.py order.xlsx payment.xlsx --json --quiet` 验证 JSON 输出格式

**预计时间**: 1.5 小时

**依赖**: 无

---

## Task 2: 更新 README.md CLI 章节

**描述**: 更新 README.md 的 `### CLI Mode` 章节（L52-87）和 `### AI Agent / Automation Mode` 章节（L88-142），确保与 AGENTS.md 和 cli.py 代码一致。

**详细内容**:

1. **检查参数表**（L77-87）:
   - 确认参数列表完整（8 个参数）
   - 确认参数说明与 AGENTS.md 一致
   - 确认默认值描述准确

2. **检查退出码表**（L104-113）:
   - 确认退出码数值与 cli.py L18-23 一致
   - 确认语义说明准确

3. **检查 JSON 格式示例**（L115-142）:
   - 确认成功信封字段与 cli.py L44-48 一致
   - 确认失败信封字段与 cli.py L35-42 一致
   - 确认 error.code 可能值与实际代码匹配

4. **补充缺失内容**（如有）:
   - stdout/stderr 分离规则
   - Agent 推荐调用方式
   - 常见错误场景

**验证步骤**:
1. 对比 README.md 和 AGENTS.md 中的参数表，确认无矛盾
2. 对比 README.md 和 cli.py 中的退出码定义，确认一致
3. 对比 README.md 和 AGENTS.md 中的 JSON 格式示例，确认一致
4. 检查 README.md 中的命令示例是否可运行

**预计时间**: 30 分钟

**依赖**: Task 1 完成后进行（以 AGENTS.md 为参考源）

---

## Task 3: 更新 documents/USAGE_EXAMPLES.md

**描述**: 检查并更新 documents/USAGE_EXAMPLES.md 的 `## CLI Mode` 章节（L42-91）和 `## AI Agent / Automation Mode` 章节（L94-194），确保与 AGENTS.md 和 cli.py 一致。

**详细内容**:

1. **检查参数表**（L81-91）:
   - 确认参数列表完整（中文说明准确）
   - 确认与 AGENTS.md 英文参数表语义对应

2. **检查退出码表**（L140-148）:
   - 确认退出码数值与 cli.py 一致
   - 确认中文说明准确

3. **检查 JSON 格式示例**（L104-137）:
   - 确认成功/失败信封结构与 cli.py 一致
   - 确认中文注释准确

4. **检查 Python 集成示例**（L172-194）:
   - 确认代码可运行
   - 确认与 AGENTS.md 推荐用法一致

5. **补充缺失内容**（如有）:
   - stdout/stderr 分离规则的中文说明
   - 常见错误场景的中文说明

**验证步骤**:
1. 对比 USAGE_EXAMPLES.md 和 AGENTS.md 中的参数表，确认中英文对应准确
2. 对比 USAGE_EXAMPLES.md 和 cli.py 中的退出码定义，确认一致
3. 对比 USAGE_EXAMPLES.md 和 AGENTS.md 中的 JSON 格式示例，确认结构一致
4. 检查 Python 集成示例代码的语法正确性

**预计时间**: 30 分钟

**依赖**: Task 1 完成后进行（以 AGENTS.md 为参考源）

---

## Task 4: 交叉验证文档一致性

**描述**: 对比所有文档中的 CLI 信息，确保无遗漏和矛盾。与 cli.py 代码进行最终验证。

**详细内容**:

1. **参数一致性检查**:
   - AGENTS.md 参数表 ↔ README.md 参数表 ↔ USAGE_EXAMPLES.md 参数表
   - 确认参数集合相同（8 个参数）
   - 确认默认值说明一致
   - 确认中英文说明准确对应

2. **退出码一致性检查**:
   - AGENTS.md 退出码表 ↔ README.md 退出码表 ↔ USAGE_EXAMPLES.md 退出码表 ↔ cli.py L18-23
   - 确认退出码数值相同（0, 1, 2, 3, 4）
   - 确认语义说明一致

3. **JSON 格式一致性检查**:
   - AGENTS.md JSON 示例 ↔ README.md JSON 示例 ↔ USAGE_EXAMPLES.md JSON 示例 ↔ cli.py L26-59
   - 确认字段名和嵌套结构相同
   - 确认 error.code 可能值集合一致

4. **工作流一致性检查**:
   - AGENTS.md 工作流说明 ↔ README.md 工作流说明 ↔ USAGE_EXAMPLES.md 工作流说明 ↔ cli.py L166-221
   - 确认销售报表两阶段流程描述一致
   - 确认参数组合说明一致

5. **代码-文档一致性检查**:
   - cli.py argparse 定义（L64-115）↔ 文档参数表
   - cli.py 退出码常量（L18-23）↔ 文档退出码表
   - cli.py `output_result()` 函数（L26-59）↔ 文档 JSON 格式
   - cli.py `main_cli()` 函数（L166-221）↔ 文档工作流说明

**验证步骤**:
1. 使用文本对比工具（如 diff）检查三份文档的 CLI 章节
2. 手动运行以下命令验证文档准确性:
   ```bash
   # 测试基本用法
   python cli.py order.xlsx payment.xlsx --json --quiet
   
   # 测试文件不存在错误
   python cli.py nonexistent.xlsx payment.xlsx --json --quiet
   echo $?  # 应输出 3
   
   # 测试销售报表工作流
   python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports --json --quiet
   ```
3. 对比实际输出与文档示例，确认一致
4. 运行 `python cli.py --help` 对比帮助信息与文档参数表

**预计时间**: 30 分钟

**依赖**: Task 1、Task 2、Task 3 全部完成

---

## 依赖关系图

```
Task 1 (AGENTS.md 新增章节)
   ├──→ Task 2 (README.md 更新)
   ├──→ Task 3 (USAGE_EXAMPLES.md 更新)
   └──→ Task 4 (交叉验证)
```

- Task 1 是基础，必须先完成
- Task 2 和 Task 3 可以并行执行
- Task 4 必须在 Task 1-3 全部完成后执行

## 总预计时间

- Task 1: 1.5 小时
- Task 2: 0.5 小时
- Task 3: 0.5 小时
- Task 4: 0.5 小时

**总计**: 3 小时

## 验证清单

- [x] AGENTS.md 新增章节插入位置正确（在 BUILD 之后、WHERE TO LOOK 之前）
- [x] AGENTS.md 包含所有 8 个 CLI 参数的完整说明
- [x] AGENTS.md 包含成功和失败两种 JSON 格式示例
- [x] AGENTS.md 包含完整的 5 种退出码表（0/1/2/3/4）
- [x] AGENTS.md 包含 `--json --quiet` 推荐用法和 stdout/stderr 分离说明
- [x] AGENTS.md 包含销售报表工作流的两阶段说明
- [x] README.md CLI 章节与 AGENTS.md 参数列表一致
- [x] README.md 退出码表与 cli.py L18-23 定义一致
- [x] documents/USAGE_EXAMPLES.md 中文说明准确无误
- [x] 所有文档中的 JSON 格式示例与 cli.py L26-59 `output_result()` 函数输出一致
- [x] 所有文档中的参数表与 cli.py L64-115 argparse 定义一致
- [x] 所有文档中的工作流说明与 cli.py L166-221 `main_cli()` 函数逻辑一致
- [x] 运行 `python cli.py --help` 输出与文档参数表匹配
- [x] 运行 `python cli.py order.xlsx payment.xlsx --json --quiet` 输出与文档 JSON 示例匹配（结构验证）
- [x] 运行 `python cli.py nonexistent.xlsx payment.xlsx --json --quiet; echo $?` 输出退出码 3
