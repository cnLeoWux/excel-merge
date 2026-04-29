## Why

The sales report generation workflow is currently only accessible via the CLI's `--month` flag. This is inconsistent with the main matching workflow, which supports interactive, CLI, and API-based execution, making the sales report feature less accessible and discoverable for non-CLI users.

## What Changes

- Add a trigger to the interactive mode (`excel_merge.py`) to allow users to initiate the sales report workflow.
- Add a trigger to the Flask API (`excel_merge_api.py`) to allow programmatic execution of the sales report workflow.
- Unify the trigger logic to be consistent across all three entry points (interactive, CLI, API).

## Capabilities

### New Capabilities
*(none)*

### Modified Capabilities
- `sales-report`: The requirements will be updated to include triggers from the interactive and API entry points.
- `cli-input`: The interactive mode's requirements will be extended to include an option for sales report generation.
- `http-api`: The API specification will be modified to include a parameter for triggering the sales report workflow.

## Impact

- `excel_merge.py`: Will be modified to include new user prompts or UI elements for starting the sales report workflow.
- `excel_merge_api.py`: The `/merge` and/or `/merge/json` endpoints will be updated to accept a new parameter (e.g., `month`) to trigger the report generation.
- `utils.py`: May require minor modifications to ensure the `process_sales_report_workflow` function is called correctly from the new entry points.
- `cli.py`: No significant changes are expected, but it will be reviewed for consistency with the other entry points.
