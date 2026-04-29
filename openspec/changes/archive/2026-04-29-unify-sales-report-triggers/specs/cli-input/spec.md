## ADDED Requirements

### Requirement: Interactive mode sales report trigger
The interactive mode (`excel_merge.py`) SHALL provide an option to trigger the sales report workflow after file selection.

#### Scenario: User opts to generate a sales report
- **WHEN** the user successfully selects an order and payment file in interactive mode
- **AND** the system prompts "Do you want to generate a sales report? (y/n)"
- **AND** the user enters 'y'
- **THEN** the system SHALL prompt the user to "Enter the report month (e.g., 202602): "
- **AND** the `process_sales_report_workflow` SHALL be called with the provided month.

#### Scenario: User declines to generate a sales report
- **WHEN** the user successfully selects an order and payment file in interactive mode
- **AND** the system prompts "Do you want to generate a sales report? (y/n)"
- **AND** the user enters 'n'
- **THEN** the standard file processing SHALL continue without triggering the sales report workflow.

#### Scenario: Invalid month format in interactive mode
- **WHEN** the user provides an invalid month format (e.g., "2026-02" or "abc")
- **THEN** the system SHALL display an error message and prompt again for a valid `YYYYMM` format.
