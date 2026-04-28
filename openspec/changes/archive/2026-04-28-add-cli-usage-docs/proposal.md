# CLI 使用文档补充提案

## 摘要

为 AGENTS.md 添加详细的 CLI 使用参考章节，使 AI Agent 能够完整理解 CLI 模式的全部功能，包括参数、工作流、JSON 输出格式、退出码、stdout/stderr 分离规则和推荐调用方式。同时更新 README.md 和 documents/USAGE_EXAMPLES.md 以确保文档一致性。

## 动机

当前 AGENTS.md 仅在 BUILD/LINT/TEST 章节包含基础的 CLI 命令示例，缺少以下关键信息：

1. **完整参数列表**：缺少参数类型、默认值、必填/可选等详细说明
2. **JSON 输出格式**：未文档化成功/失败两种 JSON 信封格式的字段定义
3. **退出码语义**：未说明 5 种退出码（0/1/2/3/4）及其使用场景
4. **Agent 调用模式**：未明确推荐的 `--json --quiet` 组合用法和 stdout/stderr 分离规则
5. **销售报表工作流**：未说明 `--month` 触发的两阶段处理流程及参数组合
6. **常见错误场景**：未涵盖文件不存在、处理失败等典型错误处理

这些缺失导致 AI Agent 在以下场景中无法有效使用本工具：
- 通过 subprocess 以非交互方式调用 CLI
- 解析 JSON 输出以获取统计数据
- 根据退出码判断执行结果
- 帮助用户调试 CLI 命令错误
- 理解销售报表功能的完整工作流

## 目标

1. **在 AGENTS.md 新增 CLI USAGE REFERENCE 章节**，包含：
   - 完整参数表：所有参数的名称、类型、默认值和说明
   - 两种工作流示例：基本匹配工作流、销售报表工作流
   - JSON 输出格式：成功和失败信封的字段定义及示例
   - 退出码表：5 种退出码及其语义说明
   - Agent 推荐用法：`--json --quiet` 组合、stdout/stderr 分离规则
   - 常见错误场景：文件不存在、处理错误的处理方式

2. **更新 README.md CLI 章节**：
   - 补充缺失的参数说明（如与 AGENTS.md 不一致之处）
   - 确保退出码表与 cli.py 代码一致
   - 确保 JSON 格式示例与实际输出一致

3. **更新 documents/USAGE_EXAMPLES.md**：
   - 确保 CLI Mode 和 AI Agent/Automation Mode 章节与 AGENTS.md 一致
   - 补充任何缺失的示例或说明
   - 保持中文说明的准确性和完整性

## 非目标

- **不修改 CLI 代码**：cli.py 的参数、退出码、JSON 格式保持现状，仅记录已有功能
- **不改变 Flask API 文档**：excel_merge_api.py 相关文档不在本提案范围内
- **不改变交互式模式文档**：excel_merge.py 相关文档不在本提案范围内
- **不创建新文档文件**：仅更新已有的 AGENTS.md、README.md、documents/USAGE_EXAMPLES.md

## 影响分析

| 文件 | 影响 | 说明 |
|------|------|------|
| `AGENTS.md` | 修改 | 在 `## BUILD / LINT / TEST COMMANDS` 章节之后、`## WHERE TO LOOK` 章节之前插入新章节 `## CLI USAGE REFERENCE`。新增内容约 150-200 行，包含参数表、工作流示例、JSON 格式、退出码表、推荐用法和错误处理说明。对现有章节无修改。 |
| `README.md` | 修改 | 更新 `### CLI Mode` 章节（L52-87）和 `### AI Agent / Automation Mode` 章节（L88-142）。确保参数表、退出码表、JSON 格式示例与 cli.py 代码和 AGENTS.md 内容一致。不涉及函数修改。 |
| `documents/USAGE_EXAMPLES.md` | 可能修改 | 检查 `## CLI Mode` 章节（L42-91）和 `## AI Agent / Automation Mode` 章节（L94-194）。如与 AGENTS.md 或 cli.py 存在不一致，更新相应内容。不涉及函数修改。 |
| `cli.py` | 无修改 | 仅作为文档编写的参考源，不修改代码。参考内容：argparse 定义（L64-115）、退出码常量（L18-23）、JSON 输出格式（L26-59）。 |

## 实施计划

1. **Phase 1：内容审核**（30分钟）
   - 对比 cli.py 代码中的 argparse 定义与现有文档
   - 确认 JSON 输出格式、退出码定义
   - 识别文档间的不一致之处

2. **Phase 2：AGENTS.md 新增章节**（1小时）
   - 编写 CLI USAGE REFERENCE 完整内容
   - 验证章节位置正确
   - 确保与 cli.py 代码一致

3. **Phase 3：更新其他文档**（1小时）
   - 更新 README.md CLI 相关章节
   - 更新 documents/USAGE_EXAMPLES.md（如需要）
   - 确保三份文档间信息一致

4. **Phase 4：交叉验证**（30分钟）
   - 对比所有文档中的 CLI 信息
   - 与 cli.py 代码进行最终验证
   - 确认无遗漏和矛盾

## 风险与缓解

| 风险 | 影响 | 缓解措施 |
|------|------|----------|
| 文档与代码不一致 | 高 | 从 cli.py 提取参数、退出码、JSON 格式作为唯一事实来源；tasks.md 中明确验证步骤 |
| 章节位置插入错误 | 中 | 在 proposal.md 和 tasks.md 中明确指定插入位置（BUILD 之后、WHERE TO LOOK 之前） |
| JSON 格式示例过时 | 中 | 运行 `python cli.py` 测试命令，复制实际输出作为示例 |
| 中英文混用不一致 | 低 | AGENTS.md 使用英文（现有风格），README.md 中英混合，USAGE_EXAMPLES.md 使用中文 |

## 验证标准

- [ ] AGENTS.md 新增章节插入位置正确（BUILD 之后、WHERE TO LOOK 之前）
- [ ] AGENTS.md 包含所有 cli.py 中定义的 8 个参数
- [ ] AGENTS.md 包含成功和失败两种 JSON 格式示例
- [ ] AGENTS.md 包含完整的 5 种退出码表（0/1/2/3/4）
- [ ] AGENTS.md 包含 `--json --quiet` 推荐用法和 stdout/stderr 分离说明
- [ ] AGENTS.md 包含销售报表工作流的两阶段说明
- [ ] README.md CLI 章节与 AGENTS.md 参数列表一致
- [ ] README.md 退出码表与 cli.py 定义一致
- [ ] documents/USAGE_EXAMPLES.md 中文说明准确无误
- [ ] 所有文档中的 JSON 格式示例与 cli.py `output_result()` 函数输出一致
- [ ] 运行 `openspec validate add-cli-usage-docs --strict` 通过
