## 1. CLI 与 Agent 契约对齐

- [ ] 1.1 更新 `AGENTS.md` 的 CLI 参考，使默认自动化工作流成为使用位置参数 `target_month` 加 `--json --quiet` 的完整销售报表工作流。
- [ ] 1.2 文档化：Agents/Skills MUST 从文件名或对话上下文推断 `target_month`，并在推断不可靠时询问用户。
- [ ] 1.3 确保 `--match-only` 仅被文档化为显式缩减工作流，而不是缺少月份时的回退。
- [ ] 1.4 更新 `.opencode/skills/excel-merge-cli/SKILL.md` 的示例和参数表，以匹配默认完整工作流契约。
- [ ] 1.5 更新面向用户的使用文档，使其不要把 `--month` 描述为当前 `cli.py` 契约，除非明确标记为未来/替代接口。

## 2. CLI 输出与行为验证

- [ ] 2.1 Verify `cli.py order_file payment_file target_month --json --quiet` returns the `ok/data/error` envelope with `total_rows`, `matched_rows`, `match_rate`, and `marked_rows` for full workflow.
- [ ] 2.2 Verify `--match-only` returns matching statistics and is only used in tests/docs as an explicit reduced mode.
- [ ] 2.3 Verify `--mark-only` returns `total_rows` and `marked_rows` statistics.
- [ ] 2.4 Verify CLI full workflow writes back to the order file in place and does not create `report_*.xlsx` artifacts.
- [ ] 2.5 Verify missing/invalid files and processing failures still map to the documented JSON error envelope and exit codes.

## 3. HTTP API 契约对齐

- [ ] 3.1 Keep `/merge/json` documented and tested as API-specific JSON with `success`, `session_id`, `download_url`, `statistics`, and `files`.
- [ ] 3.2 Add or confirm `/merge/json` file-extension validation behavior matches the chosen contract.
- [ ] 3.3 Verify `/merge/json` with `month` persists the filtered report DataFrame under `results/` and returns `statistics.report_rows` plus a report download URL.
- [ ] 3.4 Verify `/merge/json` without `month` runs the standard matching workflow and returns no sales-report artifacts.
- [ ] 3.5 Verify `/merge` with `month` returns a downloadable report attachment when report data is produced.
- [ ] 3.6 Verify API error responses remain API-shaped and use appropriate HTTP status codes.

## 4. 测试套件更新

- [ ] 4.1 Update CLI tests to cover default full workflow using positional `target_month`.
- [ ] 4.2 Update CLI tests to cover Agent/Skill missing-month behavior through documentation/Skill assertions rather than silent matching-only fallback.
- [ ] 4.3 Update CLI tests for `--match-only` and `--mark-only` as explicit reduced modes.
- [ ] 4.4 Update API integration tests to assert the API-specific `success` response shape for `/merge/json`.
- [ ] 4.5 Update API integration tests for file validation, month/no-month behavior, download URLs, and report artifacts.
- [ ] 4.6 Remove or revise tests that expect CLI/API JSON shapes contradicting this change’s specs.

## 5. 文档与校验

- [ ] 5.1 Run `openspec validate align-cli-api-contracts --strict` and fix any change-level spec issues.
- [ ] 5.2 Run `openspec validate --all --strict` and fix any global spec issues.
- [ ] 5.3 Run the relevant pytest suite, including CLI and API integration tests.
- [ ] 5.4 Review docs and Skill examples for consistent wording around `target_month`, full workflow default, API-specific JSON, and no CLI `report_*.xlsx` output.
