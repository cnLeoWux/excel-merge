## Context

The repository now has a top-level `workflow_service.py` that wraps existing `utils.py` business logic and is used by `cli.py`, `excel_merge.py`, and `excel_merge_api.py`. This achieved the main structural seam recommended by the refactoring exploration, but verification found remaining boundary gaps: invalid months are validated in CLI but not consistently in service/API paths, some workflow errors are still classified by adapters, service error tests are mostly happy-path, and Agent/user docs still show old `--month` examples even though `cli.py` uses positional `target_month`.

This change is a hardening step before any `utils.py` split. It should make the workflow/service boundary reliable and documented, without changing matching behavior or introducing a new persistence model.

## Goals / Non-Goals

**Goals:**

- Make service operations reject invalid `target_month` / API `month` values with normalized `WorkflowError(code="usage_error")` before calling core sales-report logic.
- Ensure file-not-found, usage, and processing failures from service operations have predictable `code`, `message`, and `exit_code` fields.
- Keep CLI and API adapters responsible for transport formatting but reduce duplicated workflow failure classification where practical.
- Fix AGENTS.md and usage examples so current CLI examples use positional `target_month`, not `--month`.
- Add focused tests for error normalization, service metadata, adapter error mapping, and documentation contract drift.

**Non-Goals:**

- Do not split `utils.py` in this change.
- Do not change payment-fee matching, P-number fallback, sales-report marking, or date-window filtering behavior.
- Do not add a CLI `--month` alias.
- Do not version or replace `/merge/json` with the CLI `ok/data/error` envelope.
- Do not implement a full safe-persistence/atomic-write policy.

## Decisions

### Decision 1: Put month format validation in the service layer

Add a small service helper for `YYYYMM` validation using the same business range as the CLI (`2020-2099`, `01-12`). `run_sales_report()` and `prepare_api_merge(month=...)` should use it before calling `process_sales_report_workflow()`.

Rationale:

- CLI validation currently protects only CLI calls; API and direct service calls can still pass invalid months into lower-level workflow code.
- Service-level validation makes the workflow boundary testable and consistent.

Alternatives considered:

- Keep validation in each adapter. Rejected because it duplicates logic and leaves direct service calls inconsistent.
- Move validation into `utils.py`. Deferred because this change should harden orchestration boundaries without changing core algorithms.

### Decision 2: Preserve adapter-owned formatting, but normalize workflow errors before formatting

Adapters should continue to format CLI JSON/text and Flask HTTP responses, but they should prefer `WorkflowError` fields when workflow execution fails. Request validation that is purely transport-level (missing upload fields, empty filenames, Flask `send_file`) remains in the adapter.

Rationale:

- This preserves public response shapes while reducing duplicated classification of workflow failures.
- It respects the existing separation where CLI handles argparse/exit and Flask handles HTTP concerns.

Alternatives considered:

- Make the service return ready-to-print CLI/API dictionaries. Rejected because it would blur transport formatting with workflow orchestration.

### Decision 3: Keep API empty-report behavior explicit

The current API service helper treats an empty filtered monthly report as a processing failure. This change should either document that behavior in specs/tests or adjust it deliberately. Unless implementation work reveals a strong compatibility reason, keep current behavior and add a test so it is no longer implicit.

Rationale:

- The behavior already exists and API callers receive an error instead of an empty report file.
- Making it explicit is lower risk than changing semantics during boundary hardening.

Alternatives considered:

- Allow empty report artifacts with `report_rows: 0`. Deferred to a future API behavior change because it may affect clients.

### Decision 4: Fix docs to current CLI, not future ergonomics

AGENTS.md and `documents/USAGE_EXAMPLES.md` should show `python cli.py order.xlsx payment.xlsx 202602 --json --quiet` for full workflow. Mentions of `--month` as current CLI syntax should be removed or explicitly labeled future/alternate, but this change does not add the alias.

Rationale:

- The code and specs currently use positional `target_month`.
- Documentation drift causes Agent automation to call a non-existent flag.

Alternatives considered:

- Add `--month` now to match existing docs. Rejected because contract alignment previously chose positional `target_month` and aliasing is a separate ergonomics change.

## Risks / Trade-offs

- [Risk] Service-level month validation may produce different API status codes for invalid months. → Mitigation: map `usage_error` to HTTP 400 and assert that behavior.
- [Risk] Moving file existence classification from CLI/API into service could subtly change error messages. → Mitigation: preserve documented codes/exit statuses and only loosen exact message assertions where appropriate.
- [Risk] Documentation updates may conflict with older examples in generated knowledge text. → Mitigation: target the CLI usage sections and add text-based tests for key forbidden/current patterns.
- [Risk] Empty report behavior remains debatable. → Mitigation: document and test current behavior, leaving semantic changes for a future API contract change.

## Migration Plan

1. Add service month validation and tests for missing/invalid month.
2. Add service tests for missing order/payment files and write failures.
3. Adjust CLI/API adapters to rely on normalized workflow errors where practical while preserving transport validation and public envelopes.
4. Update AGENTS.md and USAGE_EXAMPLES.md to positional `target_month` examples and explicit reduced-mode language.
5. Add/update tests for documentation drift and adapter error mapping.
6. Run OpenSpec validation and targeted pytest suites.

Rollback strategy: revert the service, adapter, docs, and tests in one commit. No data migration or dependency change is involved.

## Open Questions

- Should a later change add a supported `--month YYYYMM` alias in addition to positional `target_month`?
- Should a later API contract allow empty report downloads instead of treating empty report data as a processing error?
- Should backup and atomic-write policy move fully into a persistence service in a separate safe-persistence change?
