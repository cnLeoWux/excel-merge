## ADDED Requirements

### Requirement: API sales report trigger
The Flask API endpoints `/merge` and `/merge/json` MUST support triggering the sales report workflow via a form parameter.

#### Scenario: Trigger sales report via /merge/json
- **WHEN** a client sends a `POST` request to `/merge/json` with valid `order_file`, `payment_file`, and a `month` form parameter (e.g., "202602")
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the JSON response `data` field MUST include `report_file` and `report_rows` keys, consistent with the CLI's JSON output for the sales report workflow.

#### Scenario: Trigger sales report via /merge
- **WHEN** a client sends a `POST` request to `/merge` with valid `order_file`, `payment_file`, and a `month` form parameter.
- **THEN** the `process_sales_report_workflow` SHALL be executed.
- **AND** the system SHALL return the generated monthly report file (`report_YYYYMM.xlsx`) as a file attachment if it was created.

#### Scenario: No month parameter provided
- **WHEN** a client sends a `POST` request to `/merge` or `/merge/json` without the `month` parameter.
- **THEN** the standard matching workflow SHALL be executed.
- **AND** the response SHALL NOT contain sales report artifacts.
