## MODIFIED Requirements

### Requirement: CLI output uses workflow service results

CLI output formatting MUST consume workflow/service result objects for `data.output_file`, `data.statistics`, and processing errors instead of recomputing shared workflow statistics in `cli.py`.

#### Scenario: Full workflow JSON from service result
- **WHEN** `cli.py` completes a full sales-report workflow through the service layer
- **THEN** CLI JSON `data.output_file` SHALL come from the service result
- **AND** CLI JSON `data.statistics` SHALL come from the service result

#### Scenario: Reduced workflow JSON from service result
- **WHEN** `cli.py` completes `--match-only` or `--mark-only` through the service layer
- **THEN** CLI JSON statistics SHALL reflect the service result for that selected mode

#### Scenario: Error mapping from service error
- **WHEN** the service layer returns or raises a normalized workflow error
- **THEN** `cli.py` SHALL map it to the documented CLI JSON error envelope and exit code

#### Scenario: Invalid month from service error
- **WHEN** the full workflow service rejects `target_month` as invalid
- **THEN** `cli.py` SHALL output JSON `error.code="usage_error"` in JSON mode
- **AND** it SHALL exit with code 2

#### Scenario: File-not-found from service error
- **WHEN** the workflow service rejects a missing input file
- **THEN** `cli.py` SHALL output JSON `error.code="file_not_found"` in JSON mode
- **AND** it SHALL exit with code 3
