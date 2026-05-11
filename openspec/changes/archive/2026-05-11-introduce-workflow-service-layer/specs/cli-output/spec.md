## ADDED Requirements

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

### Requirement: CLI adapter remains responsible for transport formatting

`cli.py` MUST remain responsible for argument parsing, interactive prompting, stdout/stderr formatting, and `sys.exit()` behavior even after workflow execution moves into the service layer.

#### Scenario: Argument parsing remains in CLI
- **WHEN** a user invokes `cli.py`
- **THEN** `cli.py` SHALL parse positional file arguments, optional `target_month`, mode flags, and output flags before calling the service layer

#### Scenario: JSON envelope remains in CLI
- **WHEN** CLI output is emitted in JSON mode
- **THEN** `cli.py` SHALL format the service result using the documented `ok/data/error` envelope
