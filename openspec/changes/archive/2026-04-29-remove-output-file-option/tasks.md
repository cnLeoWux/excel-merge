## 1. utils.py 工作流函数精简

- [x] 1.1 在 `filter_unmarked_and_generate_report()` 中移除 `output_dir` 形参，删除函数体内所有 `to_excel(...)` / `Path(output_dir)` / 报表文件命名（`report_{target_month}.xlsx`）相关代码；保留返回 `(updated_df, filtered_report_df)` 元组的语义
- [x] 1.2 在 `process_sales_report_workflow()` 中移除 `output_dir` 形参，更新内部对 `filter_unmarked_and_generate_report()` 的调用以匹配新签名
- [x] 1.3 更新两个函数的 docstring：明确"本工作流不写出任何文件；返回的 DataFrame 仅供调用方决定如何持久化"
- [x] 1.4 通过 `grep -rn "output_dir" utils.py` 确认无残留引用

## 2. cli.py 参数与流程重构

- [x] 2.1 从 `parser.add_argument` 调用列表中删除 `-o/--output` 与 `--output-dir` 两个参数注册块
- [x] 2.2 删除 `main_cli()` 中所有引用 `args.output` 的代码路径（基本分支与销售报表分支各有一处写文件分叉）
- [x] 2.3 删除 `main_cli()` 中所有引用 `args.output_dir` 的代码路径
- [x] 2.4 重构销售报表分支：调用 `process_sales_report_workflow(args.order_file, args.payment_file, args.month, verbose=verbose)`，对返回的 `updated_df` 调用 `write_result_file(updated_df, Path(args.order_file))` 就地写回；删除月报文件路径计算与"新报表文件: ..."相关的 print
- [x] 2.5 把销售报表分支的 `write_result_file` `try/except` 与基本分支统一为：异常 → `output_result(error={code: "processing_error", message: str(e)})` → `sys.exit(EXIT_PROCESSING_ERROR)`；删除 `warnings = []` 与所有 `warnings.append(...)` 调用
- [x] 2.6 在 `output_result()` 文本模式中删除 `"Report saved to:"` 分支与 `"warnings"` 遍历分支；将 `output_file` 的展示文案统一为表达"就地更新订单文件"的中文摘要
- [x] 2.7 精简 JSON 成功响应：基本分支与销售报表分支共用同一构造逻辑，`data` 仅含 `output_file`（始终等于 `args.order_file`）与 `statistics`（`total_rows`/`matched_rows`/`match_rate`）；不再写入 `report_file`、`report_rows`、`warnings` 字段
- [x] 2.8 通过 `grep -nE "output_dir|--output|args\.output" cli.py` 确认无残留引用

## 3. excel_merge_api.py 内部适配（design D3）

- [x] 3.1 更新 `excel_merge_api.py` 中两处对 `process_sales_report_workflow(...)` 的调用，移除 `output_dir=RESULT_FOLDER` 实参
- [x] 3.2 在两处 `/merge` 与 `/merge/json` 销售报表分支内，使用返回的 `report_df` 自行调用 `report_df.to_excel(result_path, index=False)` 写出 `report_{month}_{session_id}.xlsx`，保持 API 对外的下载契约不变
- [x] 3.3 启动 `python excel_merge_api.py` 并对 `/health`、`/merge`、`/merge/json` 做一次冒烟，确认无 ImportError / TypeError；销售报表请求仍能下载到 `report_*.xlsx`

## 4. 自动化测试

- [x] 4.1 在 `tests/integration/` 中新增用例：`python cli.py order.xlsx payment.xlsx -o result.xlsx` 子进程必须以退出码 2 退出，stderr 含 argparse 的 unrecognized argument 信息
- [x] 4.2 新增用例：`python cli.py order.xlsx payment.xlsx --output-dir ./out` 同样必须退出码 2
- [x] 4.3 新增用例：`python cli.py order.xlsx payment.xlsx --month 202602 --json --quiet`，断言：(a) 退出码 0；(b) stdout JSON `data` 不含 `report_file`/`report_rows`/`warnings`；(c) cwd 与项目内任何目录均无新增 `report_*.xlsx`
- [x] 4.4 新增用例：把订单文件设为只读，运行 `python cli.py order.xlsx payment.xlsx --month 202602 --json`，断言：退出码 4，stdout JSON `error.code == "processing_error"`
- [x] 4.5 在 `tests/unit/` 中新增/调整 `filter_unmarked_and_generate_report` 与 `process_sales_report_workflow` 的单测：调用后断言文件系统快照与调用前一致（无新增文件）；签名不接受 `output_dir`
- [x] 4.6 调整既有任何依赖 `output_dir` 或 `report_file` JSON 字段的测试，使其与新契约对齐
- [x] 4.7 运行 `python -m pytest` 全部通过

## 5. 文档同步

- [x] 5.1 更新 `AGENTS.md` 的 "CLI USAGE REFERENCE" 整节：参数表删除 `-o/--output`、`--output-dir`；JSON 成功示例 `data` 移除 `report_file`/`report_rows`/`warnings`；命令示例集合替换为只就地修改的形式；销售报表工作流文档改为"就地写回订单文件、不产生独立报表文件"
- [x] 5.2 更新 `documents/USAGE_EXAMPLES.md`：清除所有 `-o`、`--output-dir`、`report_*.xlsx` 引用与示例；在文档顶部新增"破坏性变更迁移"小节，提示用户先复制订单文件再运行 CLI 以获得"另存"效果
- [x] 5.3 更新 `.opencode/skills/excel-merge-cli/SKILL.md`：参数表、决策树、调用模板、JSON shape、速查卡均删除被移除参数与月报文件相关条目；重写"安全提示"段为"所有产出始终就地写回订单文件，运行前请自行备份"
- [x] 5.4 检查 `README.md`（若存在 CLI 段落）按 `agent-documentation` capability 的"文档一致性"要求同步
- [x] 5.5 全仓搜索 `grep -rnE "(--output(-dir)?|report_[0-9]+\.xlsx|report_YYYYMM)" --include="*.md"` 确认无残留

## 6. OpenSpec 存档前校验

- [x] 6.1 运行 `openspec validate remove-output-file-option --strict`，全部通过
- [x] 6.2 人工对照 `proposal.md` 的 What Changes 与 Modified Capabilities，确认 6 项任务组覆盖每一条承诺
- [x] 6.3 人工对照 `design.md` 的 6 个决策（D1-D6），确认每个决策都被任务实现或显式接受
