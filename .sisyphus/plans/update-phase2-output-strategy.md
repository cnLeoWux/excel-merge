# Update Phase 2 Output Strategy — Spec & Docs Alignment

## TL;DR

> **Quick Summary**: Align the OpenSpec `sales-report` capability spec and human-facing documentation with Phase 2's actual implemented behavior — namely that Phase 2 writes `销售报表YYYYMM` markings back to the original order file on disk in addition to producing `report_YYYYMM.xlsx`. Also correct two latent inaccuracies discovered during audit (travel-date window described as "前后 1 年" but implemented as "前一年至目标月份"; second marking pass undocumented).
>
> **Deliverables**:
> - 4 OpenSpec artifacts under `openspec/changes/update-phase2-output-strategy/`: `proposal.md`, `design.md`, `specs/sales-report/spec.md` (delta), `tasks.md`
> - 5 updated human docs: `README.md`, `USAGE.md`, `documents/ARCHITECTURE.md`, `documents/TECHNICAL_DOCS.md`, `AGENTS.md`
> - `openspec validate update-phase2-output-strategy --strict` passes with exit code 0
>
> **Estimated Effort**: Short (markdown-only; no code; no tests)
> **Parallel Execution**: YES — 3 waves
> **Critical Path**: T1 (proposal) → T3 (specs delta) → T10 (validate)

---

## Context

### Original Request
> 按照 openspec 最佳实践完善文档，我希望分析二阶段工作流，匹配结果应该和一阶段一样更新原文件，而不是新建个新文件。

### Interview Summary
**Decisions confirmed**:
- **Goal interpretation**: "Keep `report_YYYYMM.xlsx` AND ensure marking is written back to the original file." (Not "drop report file", not "extra sheet inside original".)
- **Scope**: **Spec + docs only**. Zero code changes. Zero behavior changes.
- **Change name**: `update-phase2-output-strategy` (kebab-case, scaffolded via `openspec new change`).

### Code Audit Findings (verified by reading source)
The user's worry is rooted in a **documentation gap**, not a behavior bug. The current code already does exactly what the user wants:

| Finding | Evidence |
|---|---|
| Phase 2 writes `销售报表YYYYMM` into the original DataFrame for filtered rows | `utils.py:883-888` |
| The updated DataFrame is persisted to disk (in-place by default, or to `-o`) | `cli.py:180-190`, `excel_merge.py:258-265` |
| Phase 2 also writes `report_YYYYMM.xlsx` containing only the filtered subset | `utils.py:867-875` |
| Travel-date window is **[target_year-1 / target_month, target_year / target_month]** (i.e. previous 12 months up to and including target month, **inclusive both ends**) | `utils.py:835-844` and string compare on `YYYYMM` at `utils.py:859` |

### Spec & Doc Inaccuracies to Correct (incidental scope, low risk)
While auditing the existing `sales-report` spec we found two pre-existing inaccuracies in the same area; fixing them belongs in this change because they share the same lines of the spec we are already touching:

1. `openspec/specs/sales-report/spec.md` Scenario `出行日期窗口` (line 38-41) says window is **"目标月份前后 1 年"**, e.g. `[2025-02-01, 2027-02-28]`. Implementation actually uses **only the previous year**, e.g. `[2025-02, 2026-02]` inclusive. → **Correct the spec to match implementation.**
2. The same spec contains no requirement covering the second marking pass (overwrite `销售报表账期` with `销售报表YYYYMM` on the rows that were copied into the report) or the persistence of the updated DataFrame to disk. → **Add a new requirement.**

### Scaffold State
- `openspec new change update-phase2-output-strategy` already executed.
- `openspec/changes/update-phase2-output-strategy/` exists with schema `spec-driven`.
- `openspec status --change update-phase2-output-strategy` reports: 0/4 artifacts complete; `proposal` is ready, `design` and `specs` blocked by `proposal`, `tasks` blocked by both.

---

## Work Objectives

### Core Objective
Produce a complete, validating OpenSpec change (4 artifacts) and synchronized human-facing documentation (5 files) that explicitly describe Phase 2's dual-output contract and correct the existing spec inaccuracies — without modifying any application code.

### Concrete Deliverables
- `openspec/changes/update-phase2-output-strategy/proposal.md`
- `openspec/changes/update-phase2-output-strategy/design.md`
- `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` (delta with `## MODIFIED Requirements` and `## ADDED Requirements`)
- `openspec/changes/update-phase2-output-strategy/tasks.md`
- Updated sections in: `README.md`, `USAGE.md`, `documents/ARCHITECTURE.md`, `documents/TECHNICAL_DOCS.md`, `AGENTS.md`

### Definition of Done
- [ ] `openspec validate update-phase2-output-strategy --strict` exits 0
- [ ] `openspec status --change update-phase2-output-strategy` reports 4/4 artifacts complete
- [ ] Every touched doc explicitly states Phase 2 produces both (a) updated original file and (b) `report_YYYYMM.xlsx`
- [ ] No `.py` file is modified (verified by `git diff --name-only` showing only `.md` files)
- [ ] No occurrence of "前后 1 年" remains in `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` (we replace with the correct window)

### Must Have
- Phase 2 dual-output contract is explicit in both the OpenSpec delta and every human doc that mentions the workflow.
- The travel-date window in the delta spec matches `utils.py` exactly: `start_month_str = f"{target_year-1}{target_month_num:02d}"`, `end_month_str = target_month`, inclusive on both ends, comparison on `YYYYMM` string.
- A new requirement in the delta covers (i) overwriting `销售报表账期 = 销售报表YYYYMM` on copied rows, and (ii) persisting the updated DataFrame to disk via `write_result_file()` (in-place to `order_file` or to `-o`).
- The change validates strictly.

### Must NOT Have (Guardrails)
- **NO** changes to `.py` files (`utils.py`, `cli.py`, `excel_merge.py`, `excel_merge_api.py`, `setup.py`, any `test_*.py`, `check_csv.py`, `debug_csv.py`, `verify_*.py`, `create_sample_data.py`).
- **NO** changes to CLI flags, JSON envelope schema, exit codes, or any HTTP API endpoint.
- **NO** rename or relocation of `report_YYYYMM.xlsx`. The filename, format, and contents are unchanged.
- **NO** new dependencies. **NO** changes to `requirements.txt` or `setup.py`.
- **NO** behavior change of any kind. The actual runtime output of `python cli.py order.xlsx payment.xlsx --month 202602` must be byte-identical before and after this change.
- **NO** edits to `openspec/specs/sales-report/spec.md` directly. All spec changes go into the delta file under `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md`. The main spec is updated only when the change is archived.
- **NO** scope creep into other specs (`cli-input`, `cli-output`, `core-matching`, `file-io`, `http-api`, `agent-documentation`). Even if mentions of the workflow appear there, leave them out of this change unless a doc update task explicitly lists them.
- **NO** new OpenSpec capabilities. Only `sales-report` is modified.

---

## Verification Strategy

> **ZERO HUMAN INTERVENTION** — every check is agent-executable.

### Test Decision
- **Infrastructure exists**: NO real automated test suite exists in this repo (root `test_*.py` files have no assertions per `AGENTS.md`).
- **Automated tests**: NONE. This change is markdown-only — no functional code paths to test.
- **Verification primary mechanism**: `openspec validate ... --strict` for spec correctness, plus `grep` assertions on doc content for human-facing files.

### QA Policy
Every task includes agent-executed QA scenarios using `Bash` (for `openspec`, `grep`, `git diff`) and `Read` (for content inspection). Evidence saved to `.sisyphus/evidence/task-{N}-{slug}.txt`.

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 1 (Start Immediately — OpenSpec artifacts; proposal first, then design+specs in parallel, tasks last):
├── T1: Write proposal.md                                    [quick, depends: none]
├── T2: Write design.md                                      [quick, depends: T1]
├── T3: Write specs/sales-report/spec.md delta               [quick, depends: T1]
└── T4: Write tasks.md                                       [quick, depends: T2, T3]

Wave 2 (After Wave 1 — Human docs, fully parallel):
├── T5: Update README.md (Sales Report Workflow section)             [quick, depends: T3]
├── T6: Update USAGE.md (销售报表工作流 section)                       [quick, depends: T3]
├── T7: Update documents/ARCHITECTURE.md (Phase 2 description)       [quick, depends: T3]
├── T8: Update documents/TECHNICAL_DOCS.md (sales-report notes)      [quick, depends: T3]
└── T9: Update AGENTS.md (Sales Report Workflow CLI section)         [quick, depends: T3]

Wave 3 (After Wave 2 — Validation):
└── T10: openspec validate --strict + git diff sanity check  [quick, depends: T1-T9]

Wave FINAL (After ALL implementation tasks — 4 parallel reviews):
├── F1: Plan compliance audit (oracle)
├── F2: Spec & doc quality review (unspecified-high)
├── F3: Real verification QA (unspecified-high)
└── F4: Scope fidelity check (deep)
→ Present results → Get explicit user okay

Critical Path: T1 → T3 → T5..T9 (parallel) → T10 → F1..F4 → user okay
Parallel Speedup: ~50% (Wave 2 has 5 parallel tasks)
Max Concurrent: 5 (Wave 2)
```

### Dependency Matrix

| Task | Depends On | Blocks |
|---|---|---|
| T1 (proposal) | — | T2, T3 |
| T2 (design) | T1 | T4 |
| T3 (specs delta) | T1 | T4, T5–T9 |
| T4 (tasks) | T2, T3 | T10 |
| T5–T9 (docs) | T3 | T10 |
| T10 (validate) | T1–T9 | F1–F4 |
| F1–F4 | T10 | user okay |

### Agent Dispatch Summary

- **Wave 1**: T1–T4 → `quick` + skill `openspec-apply-change`
- **Wave 2**: T5–T9 → `quick` (markdown editing only)
- **Wave 3**: T10 → `quick`
- **Wave FINAL**: F1 → `oracle`, F2 → `unspecified-high`, F3 → `unspecified-high`, F4 → `deep`

---

## TODOs

- [ ] 1. Write `proposal.md`

  **What to do**:
  - Create `openspec/changes/update-phase2-output-strategy/proposal.md` matching the OpenSpec proposal artifact template.
  - Sections: `## Why`, `## What Changes`, `## Capabilities` (with `### New Capabilities` empty and `### Modified Capabilities` listing only `sales-report`), `## Impact`.
  - **Why** must explain: Phase 2 already writes 销售报表YYYYMM markings back to the original file (utils.py:887-888, persisted at cli.py:180-190 / excel_merge.py:258-265), but the spec and docs describe Phase 2 as if the only output were `report_YYYYMM.xlsx`. We're aligning docs with reality, plus correcting two latent inaccuracies in the same area: (1) travel-date window described as "前后 1 年" but implemented as previous 12 months up to and including target month; (2) second marking pass undocumented.
  - **What Changes** must list: spec clarifies dual-output contract; spec corrects travel-date window; spec adds requirement for in-place mark-back persistence; 5 doc files updated to match. Mark explicitly **NOT BREAKING** (no behavior change).
  - **Capabilities → Modified**: only `sales-report`. New: none.
  - **Impact** must list every file touched (delta spec + 5 docs) and an explicit out-of-scope list (no .py changes; no CLI flag changes; no JSON envelope changes; no rename of `report_YYYYMM.xlsx`).
  - Keep concise (1–2 pages). Implementation details belong in `design.md`, not here.

  **Must NOT do**:
  - Do NOT include scenario language or `### Requirement:` headings here — those belong in the delta spec (T3).
  - Do NOT list capabilities other than `sales-report` under Modified.
  - Do NOT mark this as breaking — it is not.

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Single short markdown file with well-known template; no research needed because the plan already enumerates the content.
  - **Skills**: [`openspec-apply-change`]
    - `openspec-apply-change`: Provides OpenSpec artifact creation conventions, header structure, and validation rules.
  - **Skills Evaluated but Omitted**:
    - `openspec-new-change`: Already executed (scaffold exists); not needed.
    - `playwright`, `dev-browser`: No browser involvement.

  **Parallelization**:
  - **Can Run In Parallel**: NO (with T2/T3) — both T2 and T3 read this file.
  - **Parallel Group**: Wave 1, runs first
  - **Blocks**: T2, T3
  - **Blocked By**: None

  **References**:

  **Pattern References** (existing OpenSpec artifacts to mimic):
  - `openspec/changes/archive/2026-04-28-add-cli-usage-docs/proposal.md` — most recent archived proposal; copy its structure exactly (Why / What Changes / Capabilities / Impact).
  - `openspec/changes/archive/2026-03-30-optimize-cli-agent-input/proposal.md` — second reference for tone and length.

  **Spec References** (to know what we're modifying):
  - `openspec/specs/sales-report/spec.md` — the existing spec being modified. Read it fully to identify the existing Requirement names (e.g., "月度报表筛选", "端到端工作流编排") that the delta will reference.

  **Code References** (must be cited verbatim in `## Why` and `## Impact`):
  - `utils.py:883-888` — second marking pass writing `销售报表YYYYMM` into 销售报表账期 column for filtered rows.
  - `utils.py:867-875` — separate `report_YYYYMM.xlsx` generation.
  - `utils.py:835-844` — actual travel-date window logic (proves the "前后 1 年" wording is wrong).
  - `cli.py:180-190` — disk persistence of the updated DataFrame.
  - `excel_merge.py:258-265` — disk persistence in interactive mode.

  **Template Reference**:
  - `openspec instructions proposal --change update-phase2-output-strategy` — run this to fetch the canonical template if uncertain.

  **WHY Each Reference Matters**:
  - Two archived proposals: gives Sisyphus the exact tone, length, and section ordering the project uses, so the new proposal feels native.
  - Existing `sales-report/spec.md`: needed because the proposal must mention which existing Requirements are being modified by name; getting the names wrong would create a confusing audit trail when the change is archived.
  - Code line references: anchor the proposal in verifiable reality. Future readers (and Momus, if invoked) can grep these line numbers to confirm the proposal isn't fabricated.

  **Acceptance Criteria**:
  - [ ] `openspec/changes/update-phase2-output-strategy/proposal.md` exists and is non-empty.
  - [ ] File contains all four required headings: `## Why`, `## What Changes`, `## Capabilities`, `## Impact` (verified by `grep -E '^## (Why|What Changes|Capabilities|Impact)$' proposal.md | wc -l` returns `4`).
  - [ ] Modified Capabilities lists `sales-report` (verified by `grep -A2 'Modified Capabilities' proposal.md | grep -q 'sales-report'`).
  - [ ] New Capabilities is empty or absent (verified by checking no kebab-case capability is listed under `### New Capabilities`).
  - [ ] At least one of the code references (`utils.py:887-888`, `utils.py:835-844`, `cli.py:180-190`) appears in the file.
  - [ ] `openspec validate update-phase2-output-strategy --strict` does not error on the proposal artifact (other artifacts may still be missing — that's OK at this stage).

  **QA Scenarios**:

  ```
  Scenario: Proposal artifact exists with correct structure
    Tool: Bash
    Preconditions: Wave 1 task T1 has been completed
    Steps:
      1. Run: test -f openspec/changes/update-phase2-output-strategy/proposal.md && echo EXISTS
      2. Run: grep -cE '^## (Why|What Changes|Capabilities|Impact)$' openspec/changes/update-phase2-output-strategy/proposal.md
      3. Run: grep -q 'sales-report' openspec/changes/update-phase2-output-strategy/proposal.md && echo "capability cited"
    Expected Result: Step 1 prints EXISTS; Step 2 prints 4; Step 3 prints "capability cited"
    Failure Indicators: Missing file, count != 4, or sales-report not mentioned
    Evidence: .sisyphus/evidence/task-1-proposal-structure.txt

  Scenario: Proposal cites real code lines (negative — no fabricated references)
    Tool: Bash
    Preconditions: T1 done
    Steps:
      1. Extract every `utils.py:N-M` and `cli.py:N-M` reference from proposal.md.
      2. For each reference, run: sed -n 'N,Mp' <file> and confirm output is non-empty.
      3. Specifically confirm `sed -n '887,888p' utils.py` contains "销售报表" and "mark_value".
    Expected Result: All extracted line ranges resolve to non-empty content; the 887-888 range contains the mark-back code.
    Evidence: .sisyphus/evidence/task-1-proposal-refs.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-1-proposal-structure.txt` — output of grep checks
  - [ ] `task-1-proposal-refs.txt` — output of line-range verification

  **Commit**: NO (groups with final commit after F1–F4 approve)

- [ ] 2. Write `design.md`

  **What to do**:
  - Create `openspec/changes/update-phase2-output-strategy/design.md` matching the OpenSpec design artifact template.
  - Sections (per `openspec instructions design`): `## Context`, `## Goals / Non-Goals`, `## Decisions`, `## Risks / Trade-offs`, `## Migration Plan`, `## Open Questions`.
  - **Context**: this is a documentation/spec alignment change. Phase 2 implementation is correct; documentation is stale. Cite the same code references as the proposal.
  - **Goals**: (1) make Phase 2 dual-output contract explicit in spec + docs; (2) correct travel-date window wording; (3) add requirement covering second marking pass + persistence.
  - **Non-Goals**: changing any code, changing any CLI/API contract, changing the report filename or format, touching specs other than `sales-report`.
  - **Decisions** (each with rationale):
    - D1: Treat all three corrections (dual-output, window, second-pass) as one coherent change rather than three separate changes — they touch overlapping spec lines and have a single user-facing message ("Phase 2 doc accuracy"). Splitting would multiply ceremony without value.
    - D2: Use `## MODIFIED Requirements` for the existing "月度报表筛选" requirement (window correction is a modification, not addition) and `## ADDED Requirements` for the new "Phase 2 在原文件回写销售报表YYYYMM" requirement.
    - D3: Update 5 human docs in addition to the spec, because the project's `agent-documentation` capability requires `AGENTS.md` accuracy and end users read README/USAGE.
    - D4: Single commit at the end (not per-file) because the change is a single coherent alignment.
  - **Risks / Trade-offs**: low — markdown only, no behavior change. Risk: doc updates miss a surface where the workflow is described. Mitigation: T10 grep-based check across the 5 listed files and a final search for any remaining "前后 1 年" wording across the repo.
  - **Migration Plan**: not applicable — no behavior change, no breaking change, no data migration. When archived, OpenSpec will fold the delta into `openspec/specs/sales-report/spec.md` automatically.
  - **Open Questions**: none — user has confirmed scope and goal interpretation in the interview.

  **Must NOT do**:
  - Do NOT include scenario language here.
  - Do NOT introduce new alternatives that the user did not ask for (e.g., "what if we wrote to a sheet instead" — already rejected in the interview).

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Short markdown; structure is dictated by OpenSpec template.
  - **Skills**: [`openspec-apply-change`]
    - `openspec-apply-change`: ensures correct artifact ordering and template compliance.

  **Parallelization**:
  - **Can Run In Parallel**: YES with T3 (different files, both read T1)
  - **Parallel Group**: Wave 1 sub-batch (T2 + T3 in parallel after T1)
  - **Blocks**: T4
  - **Blocked By**: T1

  **References**:

  **Pattern References**:
  - `openspec/changes/archive/2026-04-28-add-cli-usage-docs/design.md` — exact structure to follow.
  - `openspec/changes/archive/2026-03-30-optimize-cli-agent-input/design.md` — second reference.

  **Template Reference**:
  - Run `openspec instructions design --change update-phase2-output-strategy` to fetch canonical template.

  **Cross-artifact Reference**:
  - `openspec/changes/update-phase2-output-strategy/proposal.md` — must align with what proposal claims; do not contradict.

  **WHY Each Reference Matters**:
  - Archived designs show the project's preferred level of detail (decisions get rationales; non-goals are explicit). Following the same shape avoids Momus rejecting on style grounds.

  **Acceptance Criteria**:
  - [ ] File exists at `openspec/changes/update-phase2-output-strategy/design.md`.
  - [ ] All required headings present: `## Context`, `## Goals / Non-Goals`, `## Decisions`, `## Risks / Trade-offs`, `## Migration Plan`, `## Open Questions` (verified by `grep -cE '^## (Context|Goals / Non-Goals|Decisions|Risks / Trade-offs|Migration Plan|Open Questions)$' design.md` returns `6`).
  - [ ] Decisions section contains at least 3 numbered/bulleted entries each with a stated rationale.
  - [ ] Non-Goals explicitly lists "no .py changes" and "no CLI/API changes".

  **QA Scenarios**:

  ```
  Scenario: design.md has all required sections
    Tool: Bash
    Preconditions: T2 done
    Steps:
      1. Run: grep -cE '^## (Context|Goals / Non-Goals|Decisions|Risks / Trade-offs|Migration Plan|Open Questions)$' openspec/changes/update-phase2-output-strategy/design.md
      2. Run: grep -A20 '## Decisions' openspec/changes/update-phase2-output-strategy/design.md | grep -cE '^- |^[0-9]+\.'
    Expected Result: Step 1 returns 6; Step 2 returns >= 3
    Evidence: .sisyphus/evidence/task-2-design-structure.txt

  Scenario: Non-Goals explicitly forbid code changes
    Tool: Bash
    Preconditions: T2 done
    Steps:
      1. Run: grep -A30 'Non-Goals' openspec/changes/update-phase2-output-strategy/design.md | grep -E '\.py|code change|CLI flag|API'
    Expected Result: at least one match referencing code/CLI/API non-goals
    Evidence: .sisyphus/evidence/task-2-design-nongoals.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-2-design-structure.txt`
  - [ ] `task-2-design-nongoals.txt`

  **Commit**: NO

- [ ] 3. Write `specs/sales-report/spec.md` delta

  **What to do**:
  - Create `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` as an OpenSpec **delta spec** (not a full spec).
  - Use OpenSpec delta operations: `## MODIFIED Requirements` and `## ADDED Requirements`. Do NOT use `## REMOVED Requirements` (we are not removing anything).
  - Under `## MODIFIED Requirements`, include the existing requirement `### Requirement: 月度报表筛选` with the **corrected** scenarios:
    - Replace Scenario `出行日期窗口` text. New wording (must match implementation `utils.py:835-844` exactly): window is `[<target_year-1><target_month_num:02d>, <target_year><target_month_num:02d>]` inclusive on both ends, comparison performed on `YYYYMM` strings. Concrete example: target `202602` → window `[202502, 202602]`. Date column auto-detected from `["出发日期", "出行日期"]`. Rows with unparseable date are excluded.
    - Keep all other scenarios under this requirement (`仅保留未标注行`, `报表文件命名`, `空结果不生成空文件`) verbatim from the existing spec.
  - Under `## ADDED Requirements`, add a new requirement covering the dual-output contract:
    - Suggested title: `### Requirement: 第二阶段在原文件回写销售报表YYYYMM 并持久化`
    - Description: After筛选 step, MUST overwrite `销售报表账期 = "销售报表YYYYMM"` for every row that was copied into the report (`utils.py:883-888`). The updated DataFrame MUST then be persisted to disk by the caller (`cli.py:180-190` / `excel_merge.py:258-265`) — in-place to the original `order_file` by default, or to `-o` when provided. This persistence MUST happen regardless of whether `report_YYYYMM.xlsx` was written.
    - Scenarios (minimum 3, all in OpenSpec WHEN/THEN format):
      - `Scenario: 在原文件回写销售报表标记` — WHEN 筛选出 N 行进入报表 THEN 这 N 行在原文件的 销售报表账期 列被写入 `销售报表YYYYMM`
      - `Scenario: 默认就地写回原订单文件` — WHEN 用户未提供 `-o` THEN 更新后的 DataFrame 被 `write_result_file()` 写回 `order_file` 路径（原地覆盖）
      - `Scenario: -o 时写到指定路径` — WHEN 用户提供 `-o new.xlsx` THEN 更新后的 DataFrame 写到 `new.xlsx`，原 `order_file` 不变
      - `Scenario: 无符合行时仍持久化原文件（payment 匹配结果与一阶段标注仍需保存）` — WHEN 筛选结果为 0 行 THEN 不生成 `report_YYYYMM.xlsx`，但 Phase 1 标注（全退/已取消）和 payment 匹配结果仍被持久化到原文件 / `-o`
  - Header convention: each `### Requirement:` block must contain at least one `#### Scenario:` (OpenSpec validation rule).

  **Must NOT do**:
  - Do NOT rewrite the entire spec — only delta operations.
  - Do NOT touch other requirements (`销售报表账期标注`, `端到端工作流编排`, `工作流 JSON 输出扩展`) unless making the minimum edit needed for the window correction.
  - Do NOT remove the existing `report_YYYYMM.xlsx` filename, location, or "空结果不生成空文件" behaviors.
  - Do NOT use the wording "前后 1 年" anywhere.
  - Do NOT modify `openspec/specs/sales-report/spec.md` directly. The delta lives only under `openspec/changes/update-phase2-output-strategy/`.

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Markdown delta with strict OpenSpec syntax; the content is fully specified above.
  - **Skills**: [`openspec-apply-change`]
    - `openspec-apply-change`: critical — provides delta syntax (MODIFIED/ADDED/REMOVED) and validation rules.

  **Parallelization**:
  - **Can Run In Parallel**: YES with T2.
  - **Parallel Group**: Wave 1 sub-batch (T2 + T3)
  - **Blocks**: T4, T5–T9 (docs reference the new contract wording)
  - **Blocked By**: T1

  **References**:

  **Pattern References** (delta spec format):
  - `openspec/changes/archive/2026-04-28-add-cli-usage-docs/specs/` — read any delta spec under this archived change to see exact `## MODIFIED Requirements` / `## ADDED Requirements` syntax.
  - `openspec/changes/archive/2026-03-30-optimize-cli-agent-input/specs/` — second reference.

  **Source-of-truth Spec** (the spec being modified):
  - `openspec/specs/sales-report/spec.md` (full current spec) — copy existing scenario text verbatim into the MODIFIED block when only one scenario is changing within a requirement.

  **Code References** (drive the new requirement):
  - `utils.py:883-888` — proves the second marking pass exists.
  - `utils.py:835-844` — proves the window is previous 12 months, not "前后 1 年".
  - `utils.py:859` — string comparison `start_month_str <= travel_date_str <= end_month_str` confirms inclusive bounds.
  - `cli.py:180-190` — disk persistence in CLI mode.
  - `excel_merge.py:258-265` — disk persistence in interactive mode.

  **Validation Reference**:
  - Run `openspec validate update-phase2-output-strategy --strict` after writing this file (T10 will also run it).

  **WHY Each Reference Matters**:
  - The archived delta specs are the only reliable example of OpenSpec delta syntax in this repo; following them avoids the most common validation failures (missing `### Requirement:` headers, missing `#### Scenario:` children).
  - The current `sales-report/spec.md` is the exact text the delta will be merged into at archive time. Copying scenario wording verbatim (where unchanged) avoids accidental wording drift.
  - The code line references are the empirical evidence backing every claim in the new requirement; without them, Momus would reject for "spec asserts behavior without reference source."

  **Acceptance Criteria**:
  - [ ] File exists at `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md`.
  - [ ] Contains both `## MODIFIED Requirements` and `## ADDED Requirements` headings (verified by `grep -c '^## \(MODIFIED\|ADDED\) Requirements$'` returns `2`).
  - [ ] Contains at least one `### Requirement:` under each delta operation (verified: `grep -c '^### Requirement:' >= 2`).
  - [ ] Every `### Requirement:` block contains at least one `#### Scenario:` child (verified by structural check).
  - [ ] Zero occurrences of the string "前后 1 年" (verified by `grep -c '前后 1 年' returns 0`).
  - [ ] The new ADDED requirement contains the strings "销售报表YYYYMM" and "原文件" (or "原订单文件").
  - [ ] `openspec validate update-phase2-output-strategy --strict` reports no errors attributable to the spec delta.

  **QA Scenarios**:

  ```
  Scenario: Delta spec uses correct OpenSpec syntax and passes strict validation
    Tool: Bash
    Preconditions: T3 done
    Steps:
      1. Run: grep -cE '^## (MODIFIED|ADDED) Requirements$' openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md
      2. Run: grep -c '^### Requirement:' openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md
      3. Run: openspec validate update-phase2-output-strategy --strict ; echo "exit=$?"
    Expected Result: Step 1 returns 2; Step 2 returns >= 2; Step 3 prints exit=0 (after T1, T2, T4 also done) OR no error pertains to the spec file
    Evidence: .sisyphus/evidence/task-3-delta-validation.txt

  Scenario: Negative — old "前后 1 年" wording is gone, new window wording is present
    Tool: Bash
    Preconditions: T3 done
    Steps:
      1. Run: grep -c '前后 1 年' openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md
      2. Run: grep -E '202502.*202602|前一年|previous 12 months' openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md
    Expected Result: Step 1 returns 0; Step 2 finds at least one match
    Evidence: .sisyphus/evidence/task-3-window-correction.txt

  Scenario: ADDED requirement covers mark-back and persistence
    Tool: Bash
    Preconditions: T3 done
    Steps:
      1. Run: awk '/^## ADDED Requirements/,/^## /' openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md | grep -E '销售报表YYYYMM|销售报表\{|原文件|原订单文件|write_result_file'
    Expected Result: at least 2 distinct matches across the patterns
    Evidence: .sisyphus/evidence/task-3-added-req.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-3-delta-validation.txt`
  - [ ] `task-3-window-correction.txt`
  - [ ] `task-3-added-req.txt`

  **Commit**: NO

- [ ] 4. Write `tasks.md`

  **What to do**:
  - Create `openspec/changes/update-phase2-output-strategy/tasks.md` enumerating implementation tasks, mirroring the work breakdown in this plan but written for the OpenSpec audience (i.e., describing what the code/spec/doc deliverables ARE, not what the planning agent did).
  - Use the OpenSpec tasks artifact format (per `openspec instructions tasks`): `## 1. Section title` headings followed by `- [ ] N.M Task description` checklist items.
  - Suggested sections:
    - `## 1. Spec Delta` — list the spec edits (MODIFIED 月度报表筛选 window scenario; ADDED 第二阶段在原文件回写 requirement with N scenarios).
    - `## 2. Documentation Updates` — list the 5 human docs and the specific section to edit in each.
    - `## 3. Validation` — `openspec validate --strict`; grep checks for dual-output language.
  - Each task line must be self-contained and verifiable.

  **Must NOT do**:
  - Do NOT include code-change tasks. There are none.
  - Do NOT cross-reference Sisyphus plan filenames (`.sisyphus/plans/...`); `tasks.md` is part of the OpenSpec artifact set and should stand alone.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`openspec-apply-change`]

  **Parallelization**:
  - **Can Run In Parallel**: NO — depends on both T2 and T3 (must reflect their content).
  - **Parallel Group**: Wave 1 final
  - **Blocks**: T10
  - **Blocked By**: T2, T3

  **References**:

  **Pattern References**:
  - `openspec/changes/archive/2026-04-28-add-cli-usage-docs/tasks.md` — exact format.
  - `openspec/changes/archive/2026-03-30-optimize-cli-agent-input/tasks.md` — second reference.

  **Cross-artifact References**:
  - `openspec/changes/update-phase2-output-strategy/proposal.md` — Impact list must match.
  - `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` — section §1 lines must match the delta operations.

  **WHY Each Reference Matters**:
  - The archived `tasks.md` files demonstrate that this project uses simple flat checklists, not nested implementation roadmaps. Matching that style keeps the artifact reviewable.

  **Acceptance Criteria**:
  - [ ] File exists at `openspec/changes/update-phase2-output-strategy/tasks.md`.
  - [ ] Contains at least 3 numbered sections (`## 1.`, `## 2.`, `## 3.`).
  - [ ] At least one task per section.
  - [ ] No task references a `.py` file by name (this change is markdown-only).

  **QA Scenarios**:

  ```
  Scenario: tasks.md has the right shape
    Tool: Bash
    Preconditions: T4 done
    Steps:
      1. Run: grep -cE '^## [0-9]+\.' openspec/changes/update-phase2-output-strategy/tasks.md
      2. Run: grep -cE '^- \[ \]' openspec/changes/update-phase2-output-strategy/tasks.md
      3. Run: grep -E '\.py' openspec/changes/update-phase2-output-strategy/tasks.md | wc -l
    Expected Result: Step 1 >= 3; Step 2 >= 3; Step 3 == 0
    Evidence: .sisyphus/evidence/task-4-tasks-shape.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-4-tasks-shape.txt`

  **Commit**: NO

- [ ] 5. Update `README.md` Sales Report Workflow section

  **What to do**:
  - Locate the existing `## Sales Report Workflow` section (`README.md:188-195`).
  - Rewrite step 4 ("Generate `report_YYYYMM.xlsx`") so it explicitly describes the **two outputs**:
    - (a) the original order file is updated in place (or written to `-o` if provided) with `销售报表YYYYMM` markings on every row that was copied into the report — exactly the same in-place semantics as Phase 1;
    - (b) `report_YYYYMM.xlsx` containing only those filtered rows is written to `--output-dir` (or current dir).
  - Correct the travel-date window wording: replace any "1-year window of the target month" with "previous 12 months up to and including the target month" (e.g., target `202602` → window `[202502, 202602]` inclusive). Use a concrete example.
  - Keep the rest of the section unchanged.
  - If a "Two-phase processing" or "Phase 2" sub-paragraph exists nearby, ensure it is consistent with the corrected window.

  **Must NOT do**:
  - Do NOT change CLI flag tables, JSON envelope description, or exit code section.
  - Do NOT remove `report_YYYYMM.xlsx` from the doc — it still exists.
  - Do NOT introduce English/Chinese mixing inconsistencies; this section is English.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [] — straight markdown editing, no special skill needed.

  **Parallelization**:
  - **Can Run In Parallel**: YES with T6, T7, T8, T9, T11
  - **Parallel Group**: Wave 2
  - **Blocks**: T10
  - **Blocked By**: T3 (need final spec wording)

  **References**:
  - `README.md:188-195` — current section to edit (start anchor: `## Sales Report Workflow`).
  - `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` — produced by T3, contains the canonical wording to mirror.
  - `utils.py:835-844` — proves the window is `[target_year-1 / target_month, target_year / target_month]` inclusive.
  - `cli.py:180-190` — proves the original file gets the in-place persistence.

  **WHY**:
  - README is the highest-traffic doc; if it says "Generate report_YYYYMM.xlsx" without mentioning the in-place mark-back, users assume Phase 2 leaves the original untouched. That's exactly the misconception the user flagged.

  **Acceptance Criteria**:
  - [ ] `## Sales Report Workflow` section in `README.md` mentions BOTH "original order file" (or "in-place" / "updated original") AND `report_YYYYMM.xlsx`.
  - [ ] Section contains an explicit window description matching `[202502, 202602]` semantics (or equivalent English).
  - [ ] No occurrence of "1-year window" or "前后 1 年" remains in `README.md`.
  - [ ] `report_YYYYMM.xlsx` is still mentioned (negative check that we didn't accidentally remove it).

  **QA Scenarios**:

  ```
  Scenario: README dual-output language present
    Tool: Bash
    Steps:
      1. awk '/^## Sales Report Workflow/,/^## /' README.md > /tmp/readme-section.txt
      2. grep -E 'in-place|updated original|original.*file' /tmp/readme-section.txt
      3. grep 'report_YYYYMM.xlsx' /tmp/readme-section.txt
    Expected Result: Both grep commands find at least one match
    Evidence: .sisyphus/evidence/task-5-readme-dual-output.txt

  Scenario: README window correction (negative)
    Tool: Bash
    Steps:
      1. grep -ic '1-year window' README.md
      2. grep -ic '前后 1 年' README.md
    Expected Result: Both return 0
    Evidence: .sisyphus/evidence/task-5-readme-window.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-5-readme-dual-output.txt`
  - [ ] `task-5-readme-window.txt`

  **Commit**: NO

- [ ] 6. Update `USAGE.md` 销售报表工作流 section

  **What to do**:
  - Locate the 销售报表工作流 section starting around `USAGE.md:32`.
  - Add or update prose that explicitly states the dual output: 一阶段在原文件标注 `全退 / 已取消`；二阶段在原文件继续标注 `销售报表YYYYMM`（与一阶段一致地原地更新或写到 `-o`），并额外生成 `report_YYYYMM.xlsx`.
  - Correct the date window wording: 目标月份往前 12 个月至目标月份（含两端，按 YYYYMM 字符串比较）。Concrete example: 目标 `202602` → 窗口 `[202502, 202602]`. Replace any "1 年窗口" / "前后 1 年" wording.

  **Must NOT do**:
  - Do NOT touch the basic matching mode section.
  - Do NOT translate the file to English.
  - Do NOT introduce contradictions with USAGE.md basic-mode description.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5, T7, T8, T9, T11
  - **Parallel Group**: Wave 2
  - **Blocks**: T10
  - **Blocked By**: T3

  **References**:
  - `USAGE.md:32-` (Chinese usage doc) — search for `销售报表工作流`, `--month`.
  - `openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md` — canonical wording.
  - `utils.py:835-844`, `utils.py:883-888`, `cli.py:180-190` — backing evidence.

  **WHY**:
  - Chinese-language users primarily consult USAGE.md. Same misconception risk as README.

  **Acceptance Criteria**:
  - [ ] 销售报表工作流 section mentions both `原文件` (or `原订单文件`) and `report_YYYYMM.xlsx`.
  - [ ] No occurrence of "前后 1 年" or "1 年窗口" remains in `USAGE.md`.
  - [ ] Concrete example with `202602 → [202502, 202602]` (or equivalent) is present.

  **QA Scenarios**:

  ```
  Scenario: USAGE dual-output language present
    Tool: Bash
    Steps:
      1. awk '/销售报表工作流/,/^# |^## /' USAGE.md > /tmp/usage-section.txt
      2. grep -E '原文件|原订单文件|原地' /tmp/usage-section.txt
      3. grep 'report_YYYYMM.xlsx' /tmp/usage-section.txt
    Expected Result: Both greps find matches
    Evidence: .sisyphus/evidence/task-6-usage-dual-output.txt

  Scenario: USAGE window correction (negative)
    Tool: Bash
    Steps:
      1. grep -c '前后 1 年\|前后1年\|1 年窗口\|1年窗口' USAGE.md
    Expected Result: 0
    Evidence: .sisyphus/evidence/task-6-usage-window.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-6-usage-dual-output.txt`
  - [ ] `task-6-usage-window.txt`

  **Commit**: NO

- [ ] 7. Update `documents/ARCHITECTURE.md` Sales Report Flow

  **What to do**:
  - Locate `### Sales Report Flow (--month)` (`documents/ARCHITECTURE.md:78-93`).
  - The diagram fragment at lines 90-93 already contains `标记为"销售报表YYYYMM"` and `输出 report_YYYYMM.xlsx`, but does not state this is **also persisted to the original file on disk** with the same semantics as Phase 1. Add explicit prose adjacent to the diagram clarifying:
    - 二阶段对原 DataFrame 写入 `销售报表YYYYMM` 后，由调用方（cli.py / excel_merge.py）通过 `write_result_file()` 持久化到原文件路径（默认就地）或 `-o` 指定路径。
    - 同时另写一份 `report_YYYYMM.xlsx` 仅含被复制的子集到 `--output-dir`。
  - If the doc anywhere mentions the date window with "1 年", correct it to "前 12 个月 至 目标月份（含两端）" with a concrete example.

  **Must NOT do**:
  - Do NOT redraw the entire ASCII diagram; minimal targeted edits only.
  - Do NOT touch the basic matching architecture section.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5, T6, T8, T9, T11
  - **Parallel Group**: Wave 2
  - **Blocks**: T10
  - **Blocked By**: T3

  **References**:
  - `documents/ARCHITECTURE.md:78-93` — Sales Report Flow section.
  - `documents/ARCHITECTURE.md:54` — `process_sales_report_workflow` row in function table.
  - `cli.py:180-190`, `excel_merge.py:258-265` — persistence call sites.

  **WHY**:
  - This doc is referenced from README's documentation index, so architectural readers land here. A diagram missing the "writes back to original" step would mislead.

  **Acceptance Criteria**:
  - [ ] `### Sales Report Flow (--month)` section mentions persistence to original file (or `-o`).
  - [ ] `report_YYYYMM.xlsx` reference is preserved.
  - [ ] No "1 年" / "前后 1 年" wording remains in this file.

  **QA Scenarios**:

  ```
  Scenario: ARCHITECTURE Phase 2 dual-output
    Tool: Bash
    Steps:
      1. awk '/Sales Report Flow/,/^## /' documents/ARCHITECTURE.md > /tmp/arch-section.txt
      2. grep -E '原文件|原地|in-place|write_result_file|持久化' /tmp/arch-section.txt
      3. grep 'report_YYYYMM' /tmp/arch-section.txt
    Expected Result: Both find matches
    Evidence: .sisyphus/evidence/task-7-arch-dual-output.txt

  Scenario: ARCHITECTURE window correction
    Tool: Bash
    Steps:
      1. grep -c '前后 1 年\|前后1年' documents/ARCHITECTURE.md
    Expected Result: 0
    Evidence: .sisyphus/evidence/task-7-arch-window.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-7-arch-dual-output.txt`
  - [ ] `task-7-arch-window.txt`

  **Commit**: NO

- [ ] 8. Update `documents/TECHNICAL_DOCS.md` Sales Report Workflow

  **What to do**:
  - Locate `## Sales Report Workflow` (`documents/TECHNICAL_DOCS.md:124`).
  - The section already lists step 5 ("在原数据中标记为'销售报表YYYYMM'") and step 6 ("将筛选出的行写入 `report_YYYYMM.xlsx`"). Add an explicit step or a note immediately after step 5 stating that the **updated DataFrame is persisted to disk by the caller (cli.py / excel_merge.py) via `write_result_file()`** — in-place to the original file by default, or to `-o`. Cite `cli.py:180-190` and `excel_merge.py:258-265`.
  - Verify (and correct if needed) the window description: "目标月份往前 12 个月（含目标月份）". The current text at `documents/TECHNICAL_DOCS.md:140-150` should be cross-checked. Replace any "1 年" / "前后 1 年" wording.

  **Must NOT do**:
  - Do NOT add Python pseudo-code; this section is descriptive.
  - Do NOT touch unrelated sections (file reading pipeline, matching algorithm, etc.).

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5, T6, T7, T9, T11
  - **Parallel Group**: Wave 2
  - **Blocks**: T10
  - **Blocked By**: T3

  **References**:
  - `documents/TECHNICAL_DOCS.md:124-160` — Sales Report Workflow section.
  - `utils.py:883-888` — second marking pass.
  - `cli.py:180-190`, `excel_merge.py:258-265` — disk persistence.

  **WHY**:
  - Technical docs are the most likely doc for an engineer debugging Phase 2. The persistence step being implicit in the caller (not in `filter_unmarked_and_generate_report` itself) is exactly the kind of detail that belongs here.

  **Acceptance Criteria**:
  - [ ] Section explicitly mentions `write_result_file()` and persistence to original file (or `-o`).
  - [ ] Section retains the `report_YYYYMM.xlsx` description.
  - [ ] No "1 年" wording remains.

  **QA Scenarios**:

  ```
  Scenario: TECHNICAL_DOCS persistence step explicit
    Tool: Bash
    Steps:
      1. awk '/## Sales Report Workflow/,/^## /' documents/TECHNICAL_DOCS.md > /tmp/tech-section.txt
      2. grep -E 'write_result_file|持久化|原文件|in-place' /tmp/tech-section.txt
      3. grep 'report_YYYYMM' /tmp/tech-section.txt
    Expected Result: Both find matches
    Evidence: .sisyphus/evidence/task-8-tech-dual-output.txt

  Scenario: TECHNICAL_DOCS window correction
    Tool: Bash
    Steps:
      1. grep -c '前后 1 年\|前后1年' documents/TECHNICAL_DOCS.md
    Expected Result: 0
    Evidence: .sisyphus/evidence/task-8-tech-window.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-8-tech-dual-output.txt`
  - [ ] `task-8-tech-window.txt`

  **Commit**: NO

- [ ] 9. Update `AGENTS.md` Sales Report Workflow CLI section

  **What to do**:
  - Locate `### Sales Report Workflow (--month)` at `AGENTS.md:88`.
  - Rewrite the Phase 2 bullet (`AGENTS.md:95`) so it states both outputs:
    - "**Phase 2 — Filter & Persist**: Filter unmarked rows whose 出行日期 falls within the previous 12 months up to and including the target month (e.g., target `202602` → window `[202502, 202602]` inclusive). For every filtered row, overwrite 销售报表账期 = `销售报表YYYYMM` in the original DataFrame. The original order file is then updated in place (or written to `-o`), AND a separate `report_YYYYMM.xlsx` containing only the filtered rows is written to `--output-dir` (or cwd)."
  - Update the "Output:" line at `AGENTS.md:105` if it understates the in-place persistence (current wording "original order file updated (or written to `-o`) + `report_YYYYMM.xlsx` in `--output-dir` (or cwd)" is actually correct — verify and leave alone if so).
  - Replace any other "1-year window" mention in this file.

  **Must NOT do**:
  - Do NOT touch other AGENTS.md sections (CODE MAP, CONVENTIONS, ANTI-PATTERNS, etc.) — they don't describe the sales-report workflow.
  - Do NOT change the parameter table (`AGENTS.md:67`); the `--month` description there is already correct.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5, T6, T7, T8, T11
  - **Parallel Group**: Wave 2
  - **Blocks**: T10
  - **Blocked By**: T3

  **References**:
  - `AGENTS.md:88-105` — Sales Report Workflow section to edit.
  - `openspec/specs/agent-documentation/spec.md:83-95` — capability requirement that AGENTS.md document this workflow with code-consistent wording. The updated AGENTS.md must continue to satisfy this spec.
  - `utils.py:835-844` — window logic.
  - `utils.py:883-888` — second marking pass.

  **WHY**:
  - AGENTS.md is consumed by AI tools. If it says "1-year window" while the code says otherwise, AI agents will produce wrong code/queries. The `agent-documentation` capability requires this file be code-consistent.

  **Acceptance Criteria**:
  - [ ] Phase 2 bullet at `AGENTS.md:95` now states (a) in-place mark-back AND (b) `report_YYYYMM.xlsx` generation.
  - [ ] Window phrasing matches "previous 12 months up to and including target month" (English) or "前 12 个月至目标月份" (Chinese).
  - [ ] No occurrence of "1-year window" remains in `AGENTS.md`.
  - [ ] `report_YYYYMM.xlsx` still mentioned.

  **QA Scenarios**:

  ```
  Scenario: AGENTS.md Phase 2 dual-output
    Tool: Bash
    Steps:
      1. awk '/### Sales Report Workflow/,/^### /' AGENTS.md > /tmp/agents-section.txt
      2. grep -E 'in-place|updated.*original|original.*file' /tmp/agents-section.txt
      3. grep 'report_YYYYMM.xlsx' /tmp/agents-section.txt
    Expected Result: Both find matches
    Evidence: .sisyphus/evidence/task-9-agents-dual-output.txt

  Scenario: AGENTS.md window correction (negative)
    Tool: Bash
    Steps:
      1. grep -c '1-year window' AGENTS.md
    Expected Result: 0
    Evidence: .sisyphus/evidence/task-9-agents-window.txt

  Scenario: agent-documentation capability still satisfied
    Tool: Bash
    Steps:
      1. grep -A2 '### Sales Report Workflow' AGENTS.md | head -5
      2. grep -E '--month|YYYYMM' AGENTS.md | head -3
    Expected Result: section heading present and --month parameter still documented
    Evidence: .sisyphus/evidence/task-9-agents-capability.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-9-agents-dual-output.txt`
  - [ ] `task-9-agents-window.txt`
  - [ ] `task-9-agents-capability.txt`

  **Commit**: NO

- [ ] 10. Update `documents/USAGE_EXAMPLES.md` 销售报表工作流 example

  **What to do**:
  - Locate `### 销售报表工作流` at `documents/USAGE_EXAMPLES.md:60`.
  - The numbered list at lines 73-77 describes what `--month` triggers. Step 4 is "生成 `report_YYYYMM.xlsx`" — add a sibling step (or an extra paragraph) that explicitly describes step 5: 二阶段在原文件回写 `销售报表YYYYMM` 并通过 `write_result_file()` 持久化到原 order 文件（默认就地）或 `-o`，与一阶段相同的语义。
  - If any "1 年窗口" / "前后 1 年" appears, correct it.

  **Must NOT do**:
  - Do NOT change the example commands at lines 64-70.
  - Do NOT touch the parameter table at line 86 or the JSON example at line 119.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5, T6, T7, T8, T9
  - **Parallel Group**: Wave 2
  - **Blocks**: T11
  - **Blocked By**: T3

  **References**:
  - `documents/USAGE_EXAMPLES.md:60-90` — section to edit.
  - `utils.py:883-888`, `cli.py:180-190` — backing evidence.

  **WHY**:
  - Usage examples are the most concrete reference users follow when running the tool. Missing the in-place mark-back step here means users ship workflows that re-process already-marked data.

  **Acceptance Criteria**:
  - [ ] Section now describes the in-place mark-back persistence in addition to `report_YYYYMM.xlsx`.
  - [ ] `report_YYYYMM.xlsx` reference preserved.
  - [ ] No "前后 1 年" wording remains in this file.

  **QA Scenarios**:

  ```
  Scenario: USAGE_EXAMPLES dual-output description
    Tool: Bash
    Steps:
      1. awk '/### 销售报表工作流/,/^### |^## /' documents/USAGE_EXAMPLES.md > /tmp/example-section.txt
      2. grep -E '原文件|原订单文件|原地|write_result_file' /tmp/example-section.txt
      3. grep 'report_YYYYMM' /tmp/example-section.txt
    Expected Result: Both find matches
    Evidence: .sisyphus/evidence/task-10-examples-dual.txt

  Scenario: USAGE_EXAMPLES window correction
    Tool: Bash
    Steps:
      1. grep -c '前后 1 年\|前后1年' documents/USAGE_EXAMPLES.md
    Expected Result: 0
    Evidence: .sisyphus/evidence/task-10-examples-window.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-10-examples-dual.txt`
  - [ ] `task-10-examples-window.txt`

  **Commit**: NO

- [ ] 11. **(Optional / minor)** Correct `openspec/project.md:49` window wording

  **What to do**:
  - At `openspec/project.md:49`, the line reads:
    `- **销售报表两阶段**：阶段一匹配 + 标注（全退/已取消）；阶段二筛选未标注 + 1 年出行日期窗口 + 生成 report_YYYYMM.xlsx。`
  - This is a single-line summary in the project context doc. It has the same window inaccuracy and omits the in-place mark-back. Replace with:
    `- **销售报表两阶段**：阶段一匹配 + 标注（全退/已取消）；阶段二在原文件回写 销售报表YYYYMM（与阶段一相同的就地/-o 写回语义）+ 按"目标月份往前 12 个月（含目标月份）"出行日期窗口生成 report_YYYYMM.xlsx。`

  **Why this is "optional"**:
  - `openspec/project.md` is a top-level OpenSpec context doc. Editing it is technically a single-line summary refresh, not a capability change. It's listed here as **OPTIONAL** because the user's brief said "complete the documentation"; this line is part of the documentation surface and inconsistent with the corrected spec. Sisyphus may include or skip this task — explicitly note the choice in the commit message either way.

  **Must NOT do**:
  - Do NOT touch any other line in `openspec/project.md` (project conventions, constraints, etc.).
  - Do NOT add new bullets.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES with T5–T10
  - **Parallel Group**: Wave 2
  - **Blocks**: T12
  - **Blocked By**: T3

  **References**:
  - `openspec/project.md:49` — the exact line.

  **Acceptance Criteria**:
  - [ ] If executed: line at `openspec/project.md:49` no longer contains "1 年出行日期窗口" and now mentions the dual-output contract.
  - [ ] If skipped: explicitly note in the final commit message "openspec/project.md not updated this round (out-of-scope minor)".
  - [ ] `git diff --stat openspec/project.md` shows at most 1 line changed (only this single bullet).

  **QA Scenarios**:

  ```
  Scenario: project.md updated (or explicitly skipped)
    Tool: Bash
    Steps:
      1. grep -c '1 年出行日期窗口' openspec/project.md
      2. git diff --stat openspec/project.md | head -3
    Expected Result: Either Step 1 returns 0 (executed) OR final commit message documents skip; Step 2 changes <= 1 line
    Evidence: .sisyphus/evidence/task-11-project-md.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-11-project-md.txt`

  **Commit**: NO

- [ ] 12. Run `openspec validate --strict` and final repo-wide grep checks

  **What to do**:
  - Run `openspec validate update-phase2-output-strategy --strict`. Confirm exit code 0.
  - Run `openspec status --change update-phase2-output-strategy`. Confirm `4/4 artifacts complete`.
  - Run repo-wide grep for any remaining occurrences of the old "前后 1 年" / "1 年窗口" / "1-year window" wording outside archived changes.
  - Run `git diff --name-only HEAD` and confirm zero `.py` files appear.
  - Save evidence files. If any check fails, do not proceed to Final Verification Wave; instead re-open the relevant Wave 2 task to fix.

  **Must NOT do**:
  - Do NOT run the actual CLI (`python cli.py ...`); behavior change is not what we're verifying.
  - Do NOT modify `openspec/specs/sales-report/spec.md` (the main spec) — only the delta.
  - Do NOT auto-archive the change. Archiving is a separate manual step after F1–F4 and user okay.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`openspec-verify-change`]
    - `openspec-verify-change`: provides exact validation invocation and interpretation of strict-mode errors.

  **Parallelization**:
  - **Can Run In Parallel**: NO — must run after T1–T11.
  - **Parallel Group**: Wave 3 (single task)
  - **Blocks**: F1, F2, F3, F4
  - **Blocked By**: T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11

  **References**:
  - All artifacts produced by T1–T11.
  - `openspec/AGENTS.md` (if it exists) — for OpenSpec workflow conventions.

  **Acceptance Criteria**:
  - [ ] `openspec validate update-phase2-output-strategy --strict` exits 0.
  - [ ] `openspec status --change update-phase2-output-strategy` reports 4/4.
  - [ ] `git diff --name-only HEAD | grep -E '\.py$'` is empty.
  - [ ] Repo-wide search for "前后 1 年" outside `openspec/changes/archive/` returns 0 matches.
  - [ ] Repo-wide search for "1-year window" outside `openspec/changes/archive/` returns 0 matches.
  - [ ] All 6 doc files (README, USAGE, AGENTS, ARCHITECTURE, TECHNICAL_DOCS, USAGE_EXAMPLES) appear in `git diff --name-only`.

  **QA Scenarios**:

  ```
  Scenario: Strict validation passes
    Tool: Bash
    Steps:
      1. openspec validate update-phase2-output-strategy --strict; echo "exit=$?"
      2. openspec status --change update-phase2-output-strategy
    Expected Result: Step 1 prints "exit=0"; Step 2 reports "4/4 artifacts complete"
    Evidence: .sisyphus/evidence/task-12-validate.txt

  Scenario: No code changes, no leftover stale wording
    Tool: Bash
    Steps:
      1. git diff --name-only HEAD | grep -cE '\.py$'
      2. grep -r '前后 1 年' --include='*.md' . | grep -v 'openspec/changes/archive/' | wc -l
      3. grep -ri '1-year window' --include='*.md' . | grep -v 'openspec/changes/archive/' | wc -l
      4. git diff --name-only HEAD | sort -u
    Expected Result: Step 1 returns 0; Step 2 returns 0; Step 3 returns 0; Step 4 lists exactly the artifact files + the 6 doc files (and optionally openspec/project.md)
    Evidence: .sisyphus/evidence/task-12-repo-grep.txt
  ```

  **Evidence to Capture**:
  - [ ] `task-12-validate.txt`
  - [ ] `task-12-repo-grep.txt`

  **Commit**: YES — single final commit after F1–F4 approve and user okay
  - Message: `docs(sales-report): align Phase 2 output contract in spec and docs`
  - Files: all artifacts + all touched .md docs
  - Pre-commit: `openspec validate update-phase2-output-strategy --strict`

---

## Final Verification Wave (MANDATORY — after ALL implementation tasks)

> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.
> **Do NOT auto-proceed after verification. Wait for user's explicit approval before marking work complete.**
> **Never mark F1-F4 as checked before getting user's okay.** Rejection or user feedback → fix → re-run → present again → wait for okay.

- [ ] F1. **Plan Compliance Audit** — `oracle`

  Read `.sisyphus/plans/update-phase2-output-strategy.md` end-to-end. For each "Must Have" entry, verify the artifact actually delivers it (read the relevant `.md` file). For each "Must NOT Have", verify the constraint was respected: run `git diff --name-only HEAD` and confirm zero `.py` files were modified, zero changes outside `openspec/changes/update-phase2-output-strategy/` and the 5 listed doc files. Confirm `report_YYYYMM.xlsx` is still mentioned (not removed) and no new CLI flags appear in any doc.

  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [ ] F2. **Spec & Doc Quality Review** — `unspecified-high`

  Run `openspec validate update-phase2-output-strategy --strict` and confirm exit 0. Run `openspec show update-phase2-output-strategy --json` and confirm structure parses. Read every artifact and check: no broken markdown, no dangling `<!-- TODO -->`, scenario format follows `### Scenario:` + bullet list with **WHEN/THEN/AND**, all file:line references match actual source. Read each updated doc and check that the dual-output contract is stated unambiguously (must mention BOTH the original-file mark-back AND the `report_YYYYMM.xlsx`).

  Output: `Validate [PASS/FAIL] | Show [PASS/FAIL] | Scenarios well-formed [N/N] | Doc clarity [N/N] | VERDICT`

- [ ] F3. **Real Verification QA** — `unspecified-high`

  Execute the QA scenarios from EVERY task in this plan exactly as written. Capture evidence files to `.sisyphus/evidence/final-qa/`. Specifically verify:
  (a) every required `grep` assertion in the doc tasks finds the expected string
  (b) every "must NOT contain" assertion confirms the negative
  (c) `openspec validate --strict` evidence is captured
  (d) `git diff --stat` shows only `.md` files changed
  (e) the runtime behavior of `cli.py --month` is unchanged (sanity: `grep -n "filtered_df.to_excel" utils.py` still returns line 875; we don't run the CLI but confirm code untouched)

  Output: `Scenarios [N/N pass] | Negative assertions [N/N] | Behavior unchanged [PASS/FAIL] | VERDICT`

- [ ] F4. **Scope Fidelity Check** — `deep`

  For each task: read "What to do", read actual diff for the files it claims to touch (`git diff <files>`). Verify 1:1 — everything in spec was written (no missing doc updates), nothing beyond spec was written (no scope creep into `cli-input`, `cli-output`, `core-matching`, `file-io`, `http-api` specs; no edits to `openspec/specs/sales-report/spec.md` directly; no `.py` edits). Detect cross-task contamination (e.g., T5 README task touching `documents/`). Flag any unaccounted file changes. Confirm `openspec/specs/sales-report/spec.md` (the main spec) is unchanged.

  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | Main spec untouched [PASS/FAIL] | VERDICT`

---

## Commit Strategy

- **Single commit at end** (after F1–F4 approve and user okay): `docs(sales-report): align Phase 2 output contract in spec and docs`
  - Files: `openspec/changes/update-phase2-output-strategy/**/*.md`, `README.md`, `USAGE.md`, `documents/ARCHITECTURE.md`, `documents/TECHNICAL_DOCS.md`, `AGENTS.md`
  - Pre-commit: `openspec validate update-phase2-output-strategy --strict`

> Rationale: this is a single coherent doc/spec change. Splitting into per-file commits would obscure the cross-cutting nature of "align all surfaces with the same contract."

---

## Success Criteria

### Verification Commands
```bash
# 1. OpenSpec change validates strictly
openspec validate update-phase2-output-strategy --strict
# Expected exit code: 0

# 2. All 4 artifacts present
openspec status --change update-phase2-output-strategy
# Expected: 4/4 artifacts complete

# 3. Zero .py files modified
git diff --name-only HEAD | grep -E '\.py$' | wc -l
# Expected: 0

# 4. All 5 docs updated
git diff --name-only HEAD | grep -E '^(README\.md|USAGE\.md|AGENTS\.md|documents/ARCHITECTURE\.md|documents/TECHNICAL_DOCS\.md)$' | sort -u | wc -l
# Expected: 5

# 5. Dual-output contract appears in every touched human doc
for f in README.md USAGE.md AGENTS.md documents/ARCHITECTURE.md documents/TECHNICAL_DOCS.md; do
  grep -l "report_YYYYMM.xlsx" "$f" && grep -lE "原(文件|订单文件)|in-place|更新原" "$f"
done
# Expected: each file referenced twice (matches both patterns)

# 6. Old "前后 1 年" wording is gone from the delta spec
grep -c "前后 1 年" openspec/changes/update-phase2-output-strategy/specs/sales-report/spec.md
# Expected: 0

# 7. Main spec untouched (only delta is edited)
git diff --name-only HEAD -- openspec/specs/sales-report/spec.md
# Expected: empty output
```

### Final Checklist
- [ ] All "Must Have" present
- [ ] All "Must NOT Have" absent
- [ ] `openspec validate ... --strict` passes
- [ ] Only `.md` files in diff
- [ ] All 4 OpenSpec artifacts present
- [ ] All 5 human docs updated with dual-output contract
- [ ] Travel-date window corrected in delta spec
- [ ] User has given explicit "okay" after F1–F4 results
