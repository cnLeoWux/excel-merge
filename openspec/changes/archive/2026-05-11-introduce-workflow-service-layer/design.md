## Context

The repository currently uses a flat Python layout where `utils.py` contains almost all core business logic: file reading, matching, sales-report marking, report filtering, date parsing, and writing. The three entry points (`cli.py`, `excel_merge.py`, and `excel_merge_api.py`) call these functions directly and each repeats parts of orchestration such as statistics construction, write-back, error mapping, and response shaping.

The preceding `align-cli-api-contracts` change establishes the intended public contracts: Agent/Skill usage defaults to the full sales-report workflow with a `target_month`, CLI writes in place and uses `ok/data/error`, and the HTTP API may keep API-specific download-oriented JSON. This change uses those contracts as constraints and introduces an application workflow/service layer without changing core matching behavior.

## Goals / Non-Goals

**Goals:**

- Add a workflow/service module that coordinates existing `utils.py` business functions behind stable application-level operations.
- Make CLI, interactive mode, and HTTP API thinner adapters that parse inputs and format outputs but do not duplicate processing orchestration.
- Centralize statistics calculation for match-only, mark-only, full sales-report, and API report responses.
- Centralize write-back/download preparation decisions while preserving current CLI and API persistence semantics.
- Normalize service-level success and error data so entry points map them to CLI JSON/text or HTTP responses consistently.
- Preserve current public behavior as specified by `cli-input`, `cli-output`, `http-api`, and `sales-report`.

**Non-Goals:**

- Do not split `utils.py` into multiple modules in this change.
- Do not change the matching algorithm, P-number/hyphen fallback behavior, or sales-report marking rules.
- Do not change the CLI argument syntax from positional `target_month` to `--month`.
- Do not change `/merge/json` to the CLI `ok/data/error` envelope.
- Do not introduce new dependencies or package layout changes.

## Decisions

### Decision 1: Add a new service module as a compatibility layer

Create a new top-level module, e.g. `workflow_service.py`, that imports and calls existing `utils.py` functions. The service layer should not move core algorithms yet; it coordinates them.

Initial service operations:

- `run_match_only(order_file, payment_file, *, verbose=False, write_back=True) -> WorkflowResult`
- `run_mark_only(order_file, *, verbose=False, write_back=True) -> WorkflowResult`
- `run_sales_report(order_file, payment_file, target_month, *, verbose=False, write_back=True) -> WorkflowResult`
- `prepare_api_merge(order_path, payment_path, original_filename, *, month=None, session_id=None, timestamp=None) -> ApiWorkflowResult`

Rationale:

- A compatibility layer reduces risk because existing `utils.py` behavior remains the source of truth.
- Entry points can converge on shared orchestration before deeper module splits.

Alternatives considered:

- Split `utils.py` first. Rejected because that combines structural movement with behavior-boundary changes.
- Build a class-heavy service framework. Rejected because this small tool benefits from simple functions and dataclasses.

### Decision 2: Use lightweight result dataclasses

Define simple dataclasses in the service module:

- `WorkflowResult`: `output_file`, `dataframe`, `statistics`, optional `report_dataframe`, optional `message`.
- `WorkflowError`: `code`, `message`, optional `exit_code`, optional original exception.
- `ApiWorkflowResult`: `result_path`, `download_name`, `download_url`, `statistics`, `files`, optional `report_dataframe`.

Rationale:

- Results become explicit and testable without adding dependencies.
- Entry points can format the same result into CLI JSON/text or HTTP JSON/file responses.

Alternatives considered:

- Return raw dictionaries everywhere. Rejected because dict shapes are harder to validate and refactor.
- Raise all errors and let entry points handle everything. Rejected because it keeps error mapping duplicated.

### Decision 3: Centralize statistics construction

Move repeated statistics calculations into service helpers:

- Match statistics: `total_rows`, `matched_rows`, `match_rate`.
- Mark statistics: `total_rows`, `marked_rows`.
- Full workflow statistics: match statistics plus `marked_rows`.
- API report statistics: full workflow fields plus `report_rows` when a report artifact is produced.

Rationale:

- The same formulas currently appear in multiple entry points.
- Centralization prevents CLI/API drift after contract alignment.

Alternatives considered:

- Leave statistics in entry points. Rejected because it is one of the main sources of duplication.

### Decision 4: Keep persistence policy explicit per adapter

The service layer should support both in-place writes and API result-file writes, but it should not hide which persistence mode is being used.

- CLI and interactive calls use in-place write-back to the order file.
- API calls write generated artifacts into `results/` and return paths/download names.
- Core `process_sales_report_workflow()` remains non-persistent; API report persistence stays adapter/service orchestration.

Rationale:

- CLI and API have intentionally different persistence contracts.
- Explicit mode flags prevent accidental API behavior from leaking into CLI.

Alternatives considered:

- Service always writes in place. Rejected because API needs downloadable server-side files.
- Service never writes. Rejected because then write-back duplication remains in entry points.

### Decision 5: Refactor entry points incrementally

Refactor one adapter at a time:

1. Add and test service functions while entry points still work.
2. Update `cli.py` to call service functions.
3. Update `excel_merge.py` to call service functions.
4. Update `excel_merge_api.py` to call API-oriented service function(s).

Rationale:

- Each step is independently testable.
- If a regression appears, the adapter causing it is clear.

Alternatives considered:

- Rewrite all entry points in one pass. Rejected because it raises regression risk.

## Risks / Trade-offs

- [Risk] The service layer may initially wrap existing messy behavior rather than simplifying it. → Mitigation: treat this as an intentional seam; deeper cleanup belongs to later changes.
- [Risk] Error mapping can become confusing if both service and adapters handle exceptions. → Mitigation: define service error codes and keep final transport formatting in adapters.
- [Risk] API file behavior may accidentally change during refactor. → Mitigation: preserve existing `results/` output and `/download/<filename>` semantics with integration tests.
- [Risk] CLI output may change due to centralized statistics. → Mitigation: assert JSON/text contract before and after refactor.
- [Risk] New module in flat layout may increase import ambiguity. → Mitigation: use absolute imports consistent with current project style and avoid package restructuring.

## Migration Plan

1. Add `workflow_service.py` with result dataclasses, statistics helpers, and service functions wrapping current `utils.py` functions.
2. Add unit tests for service statistics, match-only, mark-only, full workflow, and API result preparation using temporary files/fixtures.
3. Refactor `cli.py` to call service functions while preserving argument parsing, interactive month prompt, logging setup, JSON envelope, and exit codes.
4. Refactor `excel_merge.py` to call service functions while preserving interactive/non-interactive behavior.
5. Refactor `excel_merge_api.py` to call service functions while preserving API-specific JSON shape, result files, and attachments.
6. Run OpenSpec validation and pytest; compare CLI/API behavior against aligned contracts.

Rollback strategy: revert the service module and adapter changes together. Since this change adds no dependencies or data migrations, rollback is a git revert.

## Open Questions

- Should `auto_backup()` be called by the service layer for all in-place writes, or remain a CLI-only concern until a separate persistence-policy change?
- Should `WorkflowError` store the original exception object, or only expose serializable fields?
- Should API service helpers own filename generation, or should Flask routes continue generating session IDs and result filenames before calling the service?
