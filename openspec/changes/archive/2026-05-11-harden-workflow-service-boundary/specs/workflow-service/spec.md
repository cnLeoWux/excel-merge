## MODIFIED Requirements

### Requirement: Workflow result structures

The workflow/service layer MUST return explicit structured results rather than requiring entry points to recompute statistics or infer persistence outputs from raw DataFrames. Known workflow failures MUST be exposed as `WorkflowError` with normalized `code`, `message`, and, when relevant to CLI, `exit_code` fields.

#### Scenario: CLI-compatible workflow result
- **WHEN** a CLI or interactive entry point receives a workflow result
- **THEN** the result SHALL include `output_file` and `statistics`
- **AND** the entry point can format the result as CLI text or the CLI `ok/data/error` JSON envelope without recomputing statistics

#### Scenario: API-compatible workflow result
- **WHEN** an API route receives an API workflow result
- **THEN** the result SHALL include `result_path`, `download_name`, `download_url`, `statistics`, and `files`
- **AND** the route can format the response using the existing API-specific JSON shape or file attachment behavior

#### Scenario: Service-level error shape
- **WHEN** the service layer handles a known workflow failure
- **THEN** it SHALL expose a normalized error code and message that entry points can map to CLI exit codes or HTTP status codes

#### Scenario: Missing input file error
- **WHEN** a service operation receives an order file or payment file path that does not exist
- **THEN** it SHALL raise `WorkflowError` with `code="file_not_found"`
- **AND** `exit_code` SHALL be 3

#### Scenario: Invalid target month error
- **WHEN** a service operation receives a non-empty `target_month` or API `month` that is not valid `YYYYMM`
- **THEN** it SHALL raise `WorkflowError` with `code="usage_error"`
- **AND** `exit_code` SHALL be 2
- **AND** it SHALL NOT call the core sales-report workflow

#### Scenario: Write failure error
- **WHEN** a service operation cannot write back or persist its result file
- **THEN** it SHALL raise `WorkflowError` with `code="processing_error"`
- **AND** `exit_code` SHALL be 4

### Requirement: Centralized statistics calculation

The workflow/service layer MUST centralize statistics calculation for shared workflows so entry points do not duplicate formulas.

#### Scenario: Match statistics
- **WHEN** a matching result DataFrame is processed
- **THEN** the service SHALL calculate `total_rows`, `matched_rows`, and `match_rate`

#### Scenario: Mark statistics
- **WHEN** a marked order DataFrame is processed
- **THEN** the service SHALL calculate `total_rows` and `marked_rows`

#### Scenario: Full workflow statistics
- **WHEN** a full sales-report workflow result is processed
- **THEN** the service SHALL calculate `total_rows`, `matched_rows`, `match_rate`, and `marked_rows`

#### Scenario: API report statistics
- **WHEN** an API sales-report request produces a filtered report DataFrame
- **THEN** the service SHALL include full workflow statistics plus `report_rows` in API-facing statistics

#### Scenario: Empty API report data
- **WHEN** an API sales-report request produces an empty filtered report DataFrame under the current contract
- **THEN** the service SHALL raise `WorkflowError` with `code="processing_error"`
- **AND** it SHALL NOT return downloadable report metadata
