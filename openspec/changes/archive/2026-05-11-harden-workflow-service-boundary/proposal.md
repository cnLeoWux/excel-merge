## Why

The workflow/service layer now centralizes orchestration, but several boundary responsibilities remain split across adapters: month validation, file-not-found handling, processing-error mapping, and some documentation still describes the older `--month` CLI shape. Hardening this boundary before splitting `utils.py` reduces the risk of carrying inconsistent error and documentation behavior into later module refactors.

## What Changes

- Strengthen service-level input validation so invalid or missing `target_month`/API `month` values produce normalized `WorkflowError(code="usage_error")` before core workflow execution.
- Add focused tests for `WorkflowError` paths: missing files, invalid month, write failures, and API service metadata behavior.
- Move more workflow failure classification into the service layer so CLI/API adapters primarily format already-normalized errors.
- Fix Agent/user-facing documentation drift by replacing current CLI `--month` examples with positional `target_month` examples and clarifying that `--match-only` is an explicit reduced workflow.
- Preserve public CLI/API envelopes and persistence behavior; no matching or sales-report algorithm changes.

## Capabilities

### New Capabilities

- None.

### Modified Capabilities

- `workflow-service`: tighten service validation/error normalization, API report statistics, and service metadata expectations.
- `cli-output`: clarify CLI formatting consumes normalized service errors while preserving existing exit codes and JSON envelope.
- `http-api`: clarify API month validation and service-error-to-HTTP mapping for workflow failures.
- `agent-documentation`: align AGENTS.md and user-facing examples with positional `target_month` and full-workflow default wording.
- `automated-testing`: require service error-normalization tests and documentation contract tests for the corrected CLI examples.

## Impact

- Affected code: `workflow_service.py`, `cli.py`, `excel_merge_api.py` and possibly small adapter cleanup in `excel_merge.py` if needed.
- Affected tests: `tests/unit/test_workflow_service.py`, CLI/API integration tests, and documentation/Skill assertion tests.
- Affected docs: `AGENTS.md`, `documents/USAGE_EXAMPLES.md`, and possibly `.opencode/skills/excel-merge-cli/SKILL.md` wording.
- No new runtime dependencies and no `utils.py` module split in this change.
