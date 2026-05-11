## ADDED Requirements

### Requirement: CLI execution routes through workflow service

`cli.py` MUST route validated workflow execution through the workflow/service layer while preserving the existing CLI input contract.

#### Scenario: Full workflow route
- **WHEN** `cli.py` has valid `order_file`, `payment_file`, and `target_month` with no reduced mode flag
- **THEN** it SHALL call the full sales-report workflow service operation

#### Scenario: Match-only route
- **WHEN** `cli.py` has valid arguments and `--match-only`
- **THEN** it SHALL call the match-only workflow service operation

#### Scenario: Mark-only route
- **WHEN** `cli.py` has valid arguments and `--mark-only`
- **THEN** it SHALL call the mark-only workflow service operation

#### Scenario: Input behavior preserved
- **WHEN** `cli.py` routes execution through the service layer
- **THEN** positional argument parsing, target-month validation, interactive target-month prompting, and mode validation SHALL remain compatible with the existing CLI input requirements
