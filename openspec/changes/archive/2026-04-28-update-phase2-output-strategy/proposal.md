## Why

Currently, the Phase 2 sales report generation (monthly report) is tightly coupled with the Phase 1 matching and order file saving. If saving the original order file fails (e.g., due to file locks or permissions), the sales report is never generated. This change decouples the report generation logic to ensure that even if the order file modification fails, the monthly report is still produced if possible.

## What Changes

- Modify `process_sales_report_workflow` to separate the "save order file" step from the "generate monthly report" step.
- Ensure that an error in Phase 1 file persistence does not block Phase 2 report generation.
- Update CLI and logic to handle cases where the order file might be read-only but the report directory is writable.

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- `sales-report`: Update the workflow sequence to allow independent execution of the report generation phase regardless of the primary order file's save status.

## Impact

- `utils.py`: `process_sales_report_workflow` and related internal error handling.
- `cli.py`: Reporting of success/failure when partial completion occurs (e.g., report saved but order file failed).
- `openspec/specs/sales-report/spec.md`: Update the capability definition to reflect this robustness requirement.
