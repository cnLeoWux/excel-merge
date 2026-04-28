## Context

Currently, `process_sales_report_workflow` and `cli.py` handle the sales report workflow by first processing the files (Phase 1) and then generating the monthly report (Phase 2). However, in `cli.py`, the `write_result_file` call for the updated order file happens BEFORE the report file is even identified or reported as saved. If `write_result_file` fails (e.g., due to the order file being open in Excel), the entire process stops, and the monthly report is never generated.

## Goals / Non-Goals

**Goals:**
- Ensure the monthly report (Phase 2) is generated and saved even if updating the primary order file (Phase 1) fails.
- Provide clear error reporting in the CLI about partial success (e.g., "Report generated but order file could not be updated").
- Maintain existing matching logic and file I/O safety.

**Non-Goals:**
- Changing the core matching algorithm.
- Modifying the Flask API (unless required for consistency).
- Implementing automatic retry logic for file locks.

## Decisions

### 1. Robustness in `cli.py`
Wrap the primary order file `write_result_file` call in a `try-except` block within the `args.month` branch. If it fails, log a warning but continue to generate/report the monthly report.

**Rationale:** The monthly report is often the critical deliverable for the user, while the in-place update of the order file (Phase 1) is a secondary convenience. Failure of the secondary task shouldn't block the primary one.

### 2. Error Accumulation in CLI Output
Update the `output_result` call to handle cases where `ok` might be `True` but there were non-fatal errors during file persistence. We will add a `warnings` or `partial_errors` list to the `data` envelope in the JSON output.

### 3. Decoupling in `process_sales_report_workflow`
Although `process_sales_report_workflow` currently returns the DataFrames, it doesn't handle saving the order file (that's left to the caller like `cli.py`). However, it *does* call `filter_unmarked_and_generate_report` which handles saving the *report* file. This separation is already somewhat present, but the orchestration in `cli.py` is where the coupling exists.

## Risks / Trade-offs

- **[Risk]** Data Inconsistency → **[Mitigation]** Ensure that if the order file save fails, the user is explicitly warned that the "marked" status (全退/已取消) was not persisted to the original file, even though the report was generated based on that data.
- **[Risk]** Confusing CLI Output → **[Mitigation]** Use distinct exit codes or clear "PARTIAL SUCCESS" messages in text mode. In JSON mode, include specific error details in a new `warnings` field.
