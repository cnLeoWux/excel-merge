## 1. CLI and Agent contract alignment

- [ ] 1.1 Update `AGENTS.md` CLI reference so the default automation workflow is the full sales-report workflow using positional `target_month` plus `--json --quiet`.
- [ ] 1.2 Document that Agents/Skills must infer `target_month` from filenames or conversation context, and ask the user when inference is not reliable.
- [ ] 1.3 Ensure `--match-only` is documented as an explicit reduced workflow only, not the missing-month fallback.
- [ ] 1.4 Update `.opencode/skills/excel-merge-cli/SKILL.md` examples and argument table to match the default full-workflow contract.
- [ ] 1.5 Update user-facing usage docs so they do not describe `--month` as the current `cli.py` contract unless clearly marked as a future/alternate interface.

## 2. CLI output and behavior verification

- [ ] 2.1 Verify `cli.py order_file payment_file target_month --json --quiet` returns the `ok/data/error` envelope with `total_rows`, `matched_rows`, `match_rate`, and `marked_rows` for full workflow.
- [ ] 2.2 Verify `--match-only` returns matching statistics and is only used in tests/docs as an explicit reduced mode.
- [ ] 2.3 Verify `--mark-only` returns `total_rows` and `marked_rows` statistics.
- [ ] 2.4 Verify CLI full workflow writes back to the order file in place and does not create `report_*.xlsx` artifacts.
- [ ] 2.5 Verify missing/invalid files and processing failures still map to the documented JSON error envelope and exit codes.

## 3. HTTP API contract alignment

- [ ] 3.1 Keep `/merge/json` documented and tested as API-specific JSON with `success`, `session_id`, `download_url`, `statistics`, and `files`.
- [ ] 3.2 Add or confirm `/merge/json` file-extension validation behavior matches the chosen contract.
- [ ] 3.3 Verify `/merge/json` with `month` persists the filtered report DataFrame under `results/` and returns `statistics.report_rows` plus a report download URL.
- [ ] 3.4 Verify `/merge/json` without `month` runs the standard matching workflow and returns no sales-report artifacts.
- [ ] 3.5 Verify `/merge` with `month` returns a downloadable report attachment when report data is produced.
- [ ] 3.6 Verify API error responses remain API-shaped and use appropriate HTTP status codes.

## 4. Test suite updates

- [ ] 4.1 Update CLI tests to cover default full workflow using positional `target_month`.
- [ ] 4.2 Update CLI tests to cover Agent/Skill missing-month behavior through documentation/Skill assertions rather than silent matching-only fallback.
- [ ] 4.3 Update CLI tests for `--match-only` and `--mark-only` as explicit reduced modes.
- [ ] 4.4 Update API integration tests to assert the API-specific `success` response shape for `/merge/json`.
- [ ] 4.5 Update API integration tests for file validation, month/no-month behavior, download URLs, and report artifacts.
- [ ] 4.6 Remove or revise tests that expect CLI/API JSON shapes contradicting this change’s specs.

## 5. Documentation and validation

- [ ] 5.1 Run `openspec validate align-cli-api-contracts --strict` and fix any change-level spec issues.
- [ ] 5.2 Run `openspec validate --all --strict` and fix any global spec issues.
- [ ] 5.3 Run the relevant pytest suite, including CLI and API integration tests.
- [ ] 5.4 Review docs and Skill examples for consistent wording around `target_month`, full workflow default, API-specific JSON, and no CLI `report_*.xlsx` output.
