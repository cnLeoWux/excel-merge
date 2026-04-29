## Context

CLI 当前的 `cli.py` 接受 `-o/--output`、`--month/--output-dir` 等参数，并通过 `utils.py` 中的 `process_sales_report_workflow` → `filter_unmarked_and_generate_report` 流水线，在两种"产出渠道"之间分叉：

1. **就地写入**：`write_result_file(updated_df, order_file)` 把匹配/标记结果写回原始订单文件。
2. **新文件产出**：基本流程下用 `-o` 指向新路径；销售报表流程下 `filter_unmarked_and_generate_report` 自身在 `output_dir` 下写出 `report_YYYYMM.xlsx`，CLI 仅在 JSON 输出中通报路径。

提案要求把这两条产出统一到第 1 条：所有合并/标记结果只能就地写回订单文件，月报数据不再落盘。这是一次跨 CLI、`utils` 工作流函数、JSON 契约、文档与 OpenSpec 四个 capability spec 的破坏性精简。

约束：
- HTTP API（`excel_merge_api.py`）目前依赖 `process_sales_report_workflow(..., output_dir=RESULT_FOLDER)` 并在 `RESULT_FOLDER` 中查找/下载 `report_*.xlsx`。提案明确 HTTP API 不在变更范围内，但工作流函数的签名/行为要变，必须显式划清边界。
- 退出码集合（0/1/2/3/4）保持不变，但"部分成功 + warnings"语义被删除。
- 项目无既有自动化测试覆盖此路径，所有验证靠新增 / 调整 `tests/` 中的样例。

## Goals / Non-Goals

**Goals:**
- CLI 只暴露唯一的产出渠道：就地修改订单文件。
- 销售报表工作流的"账期标记"职责保留，但不再产出 `report_YYYYMM.xlsx`；月报 DataFrame 仅作为内存中的中间产物存在。
- argparse 自动拒绝 `-o`、`--output-dir`，让破坏性以最清晰的方式暴露（退出码 2，stderr 说明）。
- JSON 输出契约被简化：成功响应的 `data` 不再含 `report_file` / `report_rows` / 月报相关 `warnings`。
- 写订单文件失败 = 整个调用失败（退出码 4 / `processing_error`），不再有"月报已生成但订单文件未更新"的部分成功路径。
- 四个相关 capability spec（`cli-input`、`cli-output`、`sales-report`、`agent-documentation`）通过 delta 同步更新。

**Non-Goals:**
- **不**修改 `excel_merge_api.py` 的对外契约（`/merge`、`/merge/json`、`/download/<file>` 路由不动）。
- **不**修改匹配算法、列识别规则、CSV 编码回退链等核心逻辑。
- **不**改变交互模式 `excel_merge.py` 的行为。
- **不**改变退出码常量或新增退出码。
- **不**为废弃参数提供"软弃用 + 警告"的过渡期；直接 BREAKING（提案已确认）。

## Decisions

### D1: 完全移除 `-o/--output` 与 `--output-dir`，不做软弃用

argparse 不再注册这两个参数。任何传入立即触发 argparse 的 unrecognized arguments 错误（退出码 2，错误消息走 stderr）。

**备选方案**:
- (a) 保留参数但忽略并打印 deprecation warning。
- (b) 保留参数，传入即报错并以退出码 4 失败。

**为何选当前方案**: 用户明确选择"完全移除"。argparse 原生错误消息已经足够清晰，且退出码 2（usage error）在语义上正好对应"参数错误"。增加自定义 deprecation 通道反而会让 JSON 契约变复杂。

### D2: 从 `process_sales_report_workflow` 与 `filter_unmarked_and_generate_report` 中移除 `output_dir` 参数与文件写出逻辑

`filter_unmarked_and_generate_report` 不再调用 `to_excel` 写 `report_YYYYMM.xlsx`；它仍然返回 `(updated_df, new_report_df)` 元组（CLI 只用第一个，第二个保留以便未来其他消费者使用，也避免 API 适配成本过高）。`process_sales_report_workflow` 同步移除 `output_dir` 形参。

**备选方案**:
- (a) 在工作流函数中保留 `output_dir`，但默认为 `None` 时跳过写文件 —— CLI 总传 `None`。
- (b) 拆出新函数 `mark_sales_report_inplace`，旧函数保留给 API 调用。

**为何选当前方案**: 提案明确"销售报表工作流不再产出新文件"，这是工作流的语义级变化，应在函数签名上反映出来，而不是靠默认值"暗示"。HTTP API 的处理见 D3。

### D3: HTTP API 适配作为本变更的"附带必要修复"，但不扩大 spec 范围

`excel_merge_api.py` 的 `/merge` 和 `/merge/json` 路由当前在销售报表分支里依赖 `process_sales_report_workflow(..., output_dir=RESULT_FOLDER)` 并把 `report_*.xlsx` 作为下载产物返回。一旦 D2 落地，API 直接 ImportError / TypeError。

本设计采取最小化适配：
- 在 API 内部，调用新签名 `process_sales_report_workflow(order_file, payment_file, target_month, verbose=...)`，自行用返回的 `report_df` 在 `RESULT_FOLDER` 写出 `report_YYYYMM_<session>.xlsx`。
- `http-api` capability spec **不**在本次变更的 Modified Capabilities 列表中（提案约束），但 `tasks.md` 必须包含 API 适配任务以保持仓库可运行。

**备选方案**:
- (a) 一并把 API 改为不再下载月报文件 —— 超出本次变更承诺范围，需要单独 proposal。
- (b) 让本次变更同时修订 `http-api` spec —— 拉大 PR 范围，与提案承诺不符。

**为何选当前方案**: 保持最小破坏面。API 行为对外不变（仍能下载月报），只是把"写文件"的责任从 utils 上移到 API 自己。CLI 那一侧才真正落实"绝不生成新文件"。

### D4: 写订单文件失败 = 退出码 4，移除 `warnings` 数组

旧实现中，销售报表分支写订单文件失败时会：append 一个 warning 字符串、继续生成月报、最终以退出码 0 + `warnings` 字段返回。删除月报后这条路径已不再有意义。

实施：在 CLI 销售报表分支与基本分支统一用 `try/except` 包裹 `write_result_file`；任何异常 → `output_result(error={code: "processing_error", message: str(e)})` → `sys.exit(EXIT_PROCESSING_ERROR)`。`data.warnings` 字段从 JSON schema 中删除。

### D5: JSON 输出 schema 精简

成功响应的 `data` 仅包含：
- `output_file`（始终等于 `order_file` 的字符串路径）
- `statistics`：`total_rows` / `matched_rows` / `match_rate`

`--month` 是否传入不再影响 `data` 形状。这让消费方解析逻辑不必分支。

### D6: 文本模式输出统一为"in-place"措辞

文本模式不再打印"Result saved to:"或"Report saved to:"。改为单行"订单文件已就地更新: <path>"（中文 / 英文表述在实施阶段定稿）。argparse 错误仍由 argparse 输出到 stderr。

## Risks / Trade-offs

- **[Risk] 现有用户/脚本依赖 `-o` 把结果写到独立路径** → Mitigation: 在 `documents/USAGE_EXAMPLES.md` 顶部增加"破坏性变更迁移"说明，提示用户先复制订单文件再运行 CLI 以获得"另存"效果。

- **[Risk] HTTP API 与 CLI 的销售报表行为出现表面不一致**（API 仍能下载月报，CLI 不再产出月报） → Mitigation: 在 D3 决策中显式接受这一不对称；后续可单独提 proposal 统一 API。在 `agent-documentation` spec 中明确 CLI 与 API 的契约边界。

- **[Risk] `report_df` 仍在 `process_sales_report_workflow` 的返回值里，可能让人误以为 CLI 还会写文件** → Mitigation: 在工作流 docstring 中明确"返回值仅供调用方决定如何持久化；本工作流自身不写文件"。

- **[Risk] 写订单文件失败现在是硬错误，可能让原本"至少拿到月报"的场景退化** → Mitigation: 接受。此场景在 D2/D4 之后已无意义（月报不再落盘）。错误消息中包含 `output_file` 路径与底层异常文本，便于排错。

- **[Trade-off] 不为废弃参数保留过渡期** → 用户已确认偏好破坏性精简而非软弃用；argparse 的 usage error 信息足够指引迁移。

- **[Trade-off] `filter_unmarked_and_generate_report` 仍返回 `report_df`** → 形参收紧但返回值保持，避免连带影响所有调用方。代价是函数名中的"generate_report"语义略有失配，需要在实施阶段考虑同步重命名（列入 `tasks.md` 的"可选清理"段）。

## Migration Plan

1. **顺序**: 先改 `utils.py`（移除 `output_dir`、停止写月报文件）→ 再改 `cli.py`（移除参数、统一错误处理、精简 JSON）→ 同步 API 内部适配（D3）→ 更新四份 spec → 更新文档（`AGENTS.md`、`USAGE_EXAMPLES.md`、SKILL）。
2. **测试**: `tests/integration/` 中新增/修改 CLI 子进程用例：
   - 传 `-o` 应退出码 2。
   - 传 `--output-dir` 应退出码 2。
   - `--month YYYYMM` 成功路径不在 cwd / 任何目录写出 `report_*.xlsx`。
   - 订单文件被锁定时退出码 4 / JSON `error.code == "processing_error"`。
3. **回滚**: 单一 PR，回滚即 `git revert`。OpenSpec 层面通过 `openspec changes archive --reverse` 不适用——直接以新 change 反向恢复行为。
4. **公告**: 在 PR 描述与 `documents/USAGE_EXAMPLES.md` 中标注 BREAKING；下次 release notes 顶部置顶。

## Open Questions

- **Q1**: 是否同时把 `filter_unmarked_and_generate_report` 重命名为更准确的 `mark_unmarked_sales_period`？倾向保留旧名以减少 diff，重命名作为单独的 follow-up change。**默认: 不重命名,留待后续。**
- **Q2**: 文本模式提示语用中文还是英文？项目内 `print` 既有中文也有英文。**默认: 沿用现状(中文为主),实施时与 `cli-output` spec 对齐。**
- **Q3**: 是否需要在 `--quiet` 模式下完全静音（包括"已就地更新"那行）？当前 `--quiet` 仅压制 INFO，不压制成功摘要。**默认: 维持现状,`--quiet` 不影响最终结果摘要。**
