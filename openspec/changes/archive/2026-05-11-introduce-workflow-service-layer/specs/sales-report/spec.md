## ADDED Requirements

### Requirement: Sales-report workflow service invocation

Entry points MUST invoke the full sales-report workflow through the workflow/service layer while preserving the existing sales-report semantics.

#### Scenario: Service delegates to existing sales-report workflow
- **WHEN** the workflow service receives a full sales-report request
- **THEN** it SHALL call the existing sales-report workflow implementation
- **AND** it SHALL preserve the returned updated order DataFrame and filtered report DataFrame

#### Scenario: CLI sales-report persistence through service
- **WHEN** CLI invokes the full sales-report service operation
- **THEN** the service SHALL write the updated order DataFrame back to the original order file
- **AND** it SHALL not persist the filtered report DataFrame as a CLI report file

#### Scenario: API sales-report persistence through service
- **WHEN** API invokes the sales-report service operation
- **THEN** the service SHALL make the filtered report DataFrame available for API result-file persistence
- **AND** it SHALL return API metadata for a downloadable report artifact
