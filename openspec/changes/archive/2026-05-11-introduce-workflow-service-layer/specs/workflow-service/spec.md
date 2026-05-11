## ADDED Requirements

### Requirement: Workflow service entry points

The system MUST provide a workflow/service layer that exposes application-level operations for matching-only, mark-only, full sales-report, and API-oriented merge workflows while delegating existing business rules to `utils.py` functions.

#### Scenario: Match-only workflow service
- **WHEN** an entry point calls the match-only service operation with an order file and payment file
- **THEN** the service SHALL run the existing payment-fee matching logic
- **AND** the service SHALL return a structured result containing the output file path, updated DataFrame, and match statistics

#### Scenario: Mark-only workflow service
- **WHEN** an entry point calls the mark-only service operation with an order file
- **THEN** the service SHALL run the existing sales-report period marking logic
- **AND** the service SHALL return a structured result containing the output file path, updated DataFrame, and marking statistics

#### Scenario: Full sales-report workflow service
- **WHEN** an entry point calls the sales-report service operation with an order file, payment file, and `target_month`
- **THEN** the service SHALL run the existing full sales-report workflow
- **AND** the service SHALL return a structured result containing the output file path, updated order DataFrame, report DataFrame, and full workflow statistics

#### Scenario: API-oriented workflow service
- **WHEN** the HTTP API calls the API-oriented workflow operation with uploaded file paths and optional `month`
- **THEN** the service SHALL produce the result path, download name, download URL, statistics, and file metadata needed by API routes
- **AND** the service SHALL preserve API-specific downloadable artifact behavior

### Requirement: Workflow result structures

The workflow/service layer MUST return explicit structured results rather than requiring entry points to recompute statistics or infer persistence outputs from raw DataFrames.

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
- **THEN** the service SHALL include `report_rows` in API-facing statistics

### Requirement: Persistence coordination

The workflow/service layer MUST coordinate write-back or API result-file persistence according to the calling adapter's contract.

#### Scenario: CLI in-place persistence
- **WHEN** the CLI calls a service operation that writes results
- **THEN** the service SHALL write the updated order DataFrame back to the original order file path
- **AND** the service SHALL NOT create a CLI `report_*.xlsx` artifact

#### Scenario: Interactive in-place persistence
- **WHEN** interactive mode calls a service operation that writes results
- **THEN** the service SHALL write the updated order DataFrame back to the selected order file path
- **AND** the service SHALL NOT create an independent report file for interactive mode

#### Scenario: API downloadable persistence
- **WHEN** the API calls a service operation that writes results for download
- **THEN** the service SHALL write the appropriate merged or report DataFrame under the configured API result directory
- **AND** the service SHALL return metadata needed for `/download/<filename>`
