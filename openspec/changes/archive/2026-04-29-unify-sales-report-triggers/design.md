## Context

Based on the proposal, the goal is to add entry points for the sales report workflow to the interactive and API modes, making it consistent with the CLI mode. Currently, only `python cli.py --month YYYYMM` can trigger this workflow.

## Goals / Non-Goals

**Goals:**
-   Implement a mechanism in the interactive mode (`excel_merge.py`) to ask the user if they want to run the sales report workflow and, if so, for which month.
-   Modify the Flask API (`excel_merge_api.py`) to accept a `month` parameter to trigger the sales report workflow.
-   Ensure the new triggers reuse the existing core logic in `utils.py.process_sales_report_workflow` without modification.

**Non-Goals:**
-   Changing the core logic of the sales report generation itself.
-   Altering the existing CLI functionality.
-   Adding new dependencies.

## Decisions

### 1. Interactive Mode (`excel_merge.py`) Integration

-   **Decision**: After the user successfully selects the order and payment files in the `main` function, a new prompt will be added to ask "Do you want to generate a sales report? (y/n)".
-   **Rationale**: This is the most natural point in the interactive flow to ask for this additional action. It doesn't interrupt the primary file selection process.
-   **Implementation**:
    -   If the user answers 'y', another prompt will ask for the target month in `YYYYMM` format.
    -   Input validation will be added to ensure the month is in the correct format.
    -   The `process_sales_report_workflow` function will be called with the file paths and the provided month.

### 2. API Mode (`excel_merge_api.py`) Integration

-   **Decision**: The `/merge` and `/merge/json` endpoints will be updated to accept an optional `month` parameter in the request's form data.
-   **Rationale**: Using form data is consistent with how the `order_file` and `payment_file` are currently handled. It avoids the need for a separate endpoint and keeps the API surface clean.
-   **Implementation**:
    -   Inside the endpoint logic, `request.form.get('month')` will be used to retrieve the value.
    -   If the `month` parameter is present and valid, the `process_sales_report_workflow` will be called instead of the standard `process_excel_files`.
    -   The JSON response for `/merge/json` will be updated to include the `report_file` and `report_rows` fields, mirroring the output of the CLI's `--json` mode when a sales report is generated.

### 3. Core Logic (`utils.py`)

-   **Decision**: No changes will be made to `process_sales_report_workflow`.
-   **Rationale**: The existing function is self-contained and already accepts all necessary parameters (`order_file`, `payment_file`, `month`, `output_dir`, `output_file_path`). The design focuses on calling this function from the new entry points correctly.

## Risks / Trade-offs

-   **[Risk]** User confusion in interactive mode if they don't understand the `YYYYMM` format.
    -   **Mitigation**: The prompt will include an example, e.g., "Enter the report month (e.g., 202602): ".
-   **[Risk]** Increased complexity in the `main` function of `excel_merge.py`.
    -   **Mitigation**: The new logic will be encapsulated in a separate helper function if it grows too large, keeping the `main` function clean.
-   **[Risk]** API users might not be aware of the new `month` parameter.
    -   **Mitigation**: The API documentation (e.g., in `README.md` or a future API spec) must be updated to reflect the new parameter.
