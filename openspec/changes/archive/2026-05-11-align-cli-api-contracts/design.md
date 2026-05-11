## Context

The project currently has three user-facing entry points over the same Excel merge and sales-report logic:

- `cli.py` for scripted/Agent usage.
- `excel_merge.py` for interactive local usage.
- `excel_merge_api.py` for HTTP upload/download usage.

The core behavior is centralized in `utils.py`, but the entry points each shape inputs, outputs, errors, and persistence differently. Recent spec alignment documented current behavior, including the important product rule that the default Agent/Skill workflow is the full sales-report workflow and therefore needs `target_month` from an explicit user answer or reliable filename/context inference.

This change focuses on contract alignment only. It should make the public behavior deliberate before a later workflow/service-layer refactor.

## Goals / Non-Goals

**Goals:**

- Make the default Agent/Skill automation path explicit: obtain `target_month` and run the full workflow unless the user explicitly requests a reduced matching-only flow.
- Keep CLI JSON predictable for automation via the existing `ok/data/error` envelope.
- Decide and document the API JSON contract so `/merge/json` is no longer in conflict with CLI docs and tests.
- Align tests and Agent-facing documentation with the selected contracts.
- Preserve in-place CLI output semantics: CLI writes the order file back in place and does not create `report_*.xlsx` files.
- Preserve API download semantics where HTTP callers receive result/report files through `results/` and `/download/<filename>`.

**Non-Goals:**

- Do not introduce the workflow/service layer in this change.
- Do not split `utils.py` or pure-ify `process_excel_files()` in this change.
- Do not change the core matching algorithm or P-number/hyphen priority in this change.
- Do not introduce new runtime dependencies.
- Do not remove existing HTTP endpoints.

## Decisions

### Decision 1: Treat full workflow as the default automation intent

For Agent/Skill usage, two provided files mean “run the complete matching and sales-report workflow” unless the user explicitly says they only want matching. Since the full workflow requires a month, the Agent/Skill must acquire `target_month` before invoking the CLI.

Rationale:

- This matches the business expectation for Feishu-style usage.
- It avoids silently under-processing files when the month is missing.
- It keeps `--match-only` as an intentional reduced workflow rather than an accidental fallback.

Alternatives considered:

- Default two-file CLI calls to matching-only. Rejected because the product expectation is full processing by default.
- Always ask for a month even when filename context is obvious. Rejected because reliable filename/context inference keeps the workflow efficient.

### Decision 2: Keep CLI invocation positional for this change

The current `cli.py` contract uses `order_file payment_file [target_month]`. This change should align docs and tests around that behavior rather than introducing `--month` immediately.

Rationale:

- It minimizes implementation scope for this contract-alignment change.
- The Skill already uses positional `target_month` successfully.
- A later CLI ergonomics change can add `--month` as a compatible alias if desired.

Alternatives considered:

- Add `--month` now and deprecate positional `target_month`. Deferred because this change already spans CLI, API, docs, and tests.

### Decision 3: Preserve CLI JSON envelope and allow workflow-specific statistics

CLI JSON should continue using:

```json
{ "ok": true, "data": { ... }, "error": null }
```

Full workflow statistics may include `marked_rows` in addition to `total_rows`, `matched_rows`, and `match_rate`. Reduced modes may expose statistics appropriate to the selected mode.

Rationale:

- The envelope is Agent-friendly and already implemented in CLI helpers.
- `marked_rows` is useful for full sales-report workflow visibility.
- Forcing every mode into an identical statistics object would either drop useful data or require meaningless fields.

Alternatives considered:

- Strictly identical statistics fields for every mode. Rejected because `--mark-only` has no meaningful match rate.

### Decision 4: Keep API response shape API-specific for now, but make it deliberate

For this change, `/merge/json` should remain API-specific with `success`, `session_id`, `download_url`, `statistics`, and `files`. Specs and tests should state that this is intentionally different from CLI JSON.

Rationale:

- HTTP callers need download URLs and file identifiers.
- This avoids breaking any existing clients that already rely on `success` and `download_url`.
- It still removes ambiguity by documenting the difference.

Alternatives considered:

- Convert API JSON to the CLI `ok/data/error` envelope. Deferred because that would be a breaking API change. It can be introduced later as a versioned endpoint or compatibility layer.

### Decision 5: Keep CLI and API persistence semantics separate

CLI and interactive mode write updated order data back to the order file and do not create independent report files. API mode may persist downloadable merged/report files under `results/` because HTTP clients need downloadable artifacts.

Rationale:

- CLI usage is file-local and in-place by contract.
- HTTP usage is request/response oriented and needs a server-side result path.
- The core sales-report workflow still does not write report files itself; API persistence remains an adapter responsibility.

Alternatives considered:

- Force API to mirror CLI in-place behavior. Rejected because uploaded files are stored in server-managed temporary paths and callers need downloads.

### Decision 6: Fix API validation asymmetry where practical

`/merge/json` should receive the same file-extension validation as `/merge` unless a compatibility reason prevents it. Error responses should remain API-shaped.

Rationale:

- Both endpoints accept the same file types and should fail early for unsupported files.
- This reduces divergent behavior between two HTTP routes.

Alternatives considered:

- Leave `/merge/json` validation weaker. Rejected because it preserves an avoidable contract gap.

## Risks / Trade-offs

- [Risk] Existing docs or tests may still assume `--month` instead of positional `target_month`. → Mitigation: update Agent docs, usage docs, and tests in the same change; leave a future change for adding `--month` alias.
- [Risk] API and CLI continue to use different JSON envelopes. → Mitigation: document the difference explicitly and keep API-specific response fields stable; consider a future versioned API envelope if needed.
- [Risk] Full workflow default may block Agent execution when month cannot be inferred. → Mitigation: Skill must ask the user instead of falling back to matching-only; this is intentional to avoid incomplete processing.
- [Risk] Updating tests may expose current implementation quirks such as `cli.py` interactive EOF cancellation. → Mitigation: cover current behavior where it remains part of contract, and isolate future behavior changes into separate changes.
- [Risk] API report files and CLI no-report-file semantics can look inconsistent. → Mitigation: specify that API persistence is adapter-level behavior while the core workflow remains in-memory.

## Migration Plan

1. Update delta specs for `cli-input`, `cli-output`, `http-api`, `agent-documentation`, `automated-testing`, and `sales-report`.
2. Update Agent-facing docs and Skill docs to describe full-workflow default and month acquisition.
3. Align tests with the selected CLI/API contracts.
4. Apply implementation changes only where required by the aligned contract, especially API file validation and documentation/test expectations.
5. Run `openspec validate --all --strict` and the pytest suite.

Rollback strategy: revert this change’s spec/doc/test/implementation updates together. Since no database migration or new dependency is expected, rollback is a git revert.

## Open Questions

- Should a future change add `--month YYYYMM` as an alias while keeping positional `target_month` for compatibility?
- Should a future API version expose the CLI-style `ok/data/error` envelope while preserving the current `/merge/json` shape?
- Should `cli.py` continue interactive prompting in JSON mode, or should Agent/Skill month acquisition fully replace that path in a later CLI cleanup?
