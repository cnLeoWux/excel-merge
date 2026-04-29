## Why

CLI 当前默认行为会就地修改原始订单文件，但同时通过 `-o/--output` 允许写到任意新文件，并通过 `--month/--output-dir` 在销售报表流程中产出额外的 `report_YYYYMM.xlsx`。这种"既能就地、又能生成新文件"的双轨设计带来三个问题：(1) 用户和自动化脚本必须维护额外的输出路径状态；(2) 销售报表工作流的"部分成功"语义（订单文件保存失败但月报已写出）让错误处理变得复杂；(3) 项目实际数据流约定是"合并结果归属订单文件本身"，多余的产物路径只是历史遗留。统一为"所有结果一律写回订单文件、绝不生成新文件"可以让 CLI 契约、退出码和 JSON 输出全部简化。

## What Changes

- **BREAKING**: 移除 `-o`/`--output` 参数。传入将由 argparse 报错并以退出码 2 退出。
- **BREAKING**: 移除 `--output-dir` 参数。传入将由 argparse 报错并以退出码 2 退出。
- **BREAKING**: 销售报表工作流（`--month YYYYMM`）不再生成 `report_YYYYMM.xlsx` 文件。月报数据仅在内存中用于驱动账期标记，所有结果（含 `销售报表账期` 列）写回原始订单文件。
- **BREAKING**: JSON 成功响应的 `data` 中移除 `report_file` 字段；`report_rows` 字段移除（或文档化为始终为 0）；月报相关的 `warnings` 项不再产生。
- 文本模式下移除"Result saved to: …"中关于独立新文件的提示，改为统一表述"Order file updated in place: <order_file>"。
- 所有合并/标记结果一律就地写回 `order_file`；写入失败直接以退出码 4 / `processing_error` 失败，不再有"部分成功"路径。
- 更新 `AGENTS.md`、`documents/USAGE_EXAMPLES.md`、`.opencode/skills/excel-merge-cli/SKILL.md` 中所有涉及 `-o`、`--output-dir`、`report_YYYYMM.xlsx` 的描述与示例。

## Capabilities

### New Capabilities

无新增能力。本次变更是对既有能力的精简。

### Modified Capabilities

- `cli-output`: 输出契约简化——文本模式只汇报"就地更新"，JSON 模式 `data` 不再含 `report_file`/`report_rows`/月报相关 `warnings`；移除"部分成功"语义，写入失败一律失败退出。
- `sales-report`: 工作流第二阶段不再产出 `report_YYYYMM.xlsx`；销售报表数据仅用于在内存中计算并将 `销售报表账期` 列写回订单文件。
- `agent-documentation`: 记录 CLI 用法的 agent 文档须移除对 `-o`、`--output-dir`、`report_YYYYMM.xlsx` 的所有引用，并明确"所有产出就地写回订单文件"的契约。

> 注：`cli-input` capability 现有 spec 实际只描述 `excel_merge.py` 交互模式与日志体系，并未定义 `-o`/`--output-dir` 参数；因此本次变更不需要 `cli-input` 的 delta。

## Impact

- **代码**: `cli.py`（`main_cli` 参数定义、销售报表分支、`output_result` 文本/ JSON 分支）；`utils.py` 中 `process_sales_report_workflow`、`filter_unmarked_and_generate_report` 不再需要写出报表文件，签名/行为相应简化。
- **API**: `excel_merge_api.py` 不在本次变更范围内；如其内部调用了 `process_sales_report_workflow`，需要回归确认其不依赖被删除的报表文件输出（仅消费 DataFrame）。
- **文档**: `AGENTS.md`（CLI Usage Reference 整节）、`documents/USAGE_EXAMPLES.md`、`.opencode/skills/excel-merge-cli/SKILL.md`、`openspec/specs/{cli-input,cli-output,sales-report,agent-documentation}/spec.md`。
- **退出码**: 维持现有 0/1/2/3/4 集合不变，但语义上"部分成功 + warnings"路径被删除。
- **依赖**: 无变化。
- **破坏性**: 任何现有自动化脚本只要传 `-o`、`--output-dir`，或依赖 `report_YYYYMM.xlsx`、JSON 中的 `report_file`/`report_rows` 字段，都会立即失败；需要在 `documents/` 中提供迁移指引。
