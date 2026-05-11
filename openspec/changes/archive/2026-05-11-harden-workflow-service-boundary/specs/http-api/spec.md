## MODIFIED Requirements

### Requirement: HTTP API routes use workflow service

HTTP API merge routes MUST use the workflow/service layer for shared processing while preserving API request validation, response shape, and downloadable artifact behavior.

#### Scenario: /merge without month uses API service
- **WHEN** `/merge` receives valid uploaded order and payment files without `month`
- **THEN** the route SHALL save uploads safely
- **AND** it SHALL call the API-oriented matching workflow service operation
- **AND** it SHALL return the service result file as an attachment

#### Scenario: /merge with month uses API service
- **WHEN** `/merge` receives valid uploaded order and payment files with `month`
- **THEN** the route SHALL call the API-oriented sales-report workflow service operation
- **AND** it SHALL return the generated report attachment when report data is produced

#### Scenario: /merge/json without month uses API service
- **WHEN** `/merge/json` receives valid uploaded order and payment files without `month`
- **THEN** the route SHALL call the API-oriented matching workflow service operation
- **AND** it SHALL format the service result using the documented API-specific `success` response shape

#### Scenario: /merge/json with month uses API service
- **WHEN** `/merge/json` receives valid uploaded order and payment files with `month`
- **THEN** the route SHALL call the API-oriented sales-report workflow service operation
- **AND** it SHALL format the service result using the documented API-specific response shape including `statistics.report_rows`

#### Scenario: API invalid month service error
- **WHEN** `/merge` or `/merge/json` receives uploaded files with an invalid `month` form parameter
- **THEN** the API-oriented workflow service SHALL reject the request with `WorkflowError(code="usage_error")`
- **AND** the route SHALL return an HTTP 400 response
- **AND** `/merge/json` SHALL return an API-shaped JSON failure response with `success=false` and `error`

### Requirement: HTTP adapter remains responsible for HTTP concerns

`excel_merge_api.py` MUST remain responsible for Flask-specific concerns such as request parsing, upload field validation, `secure_filename()`, HTTP status codes, and `send_file()` responses.

#### Scenario: Upload validation before service call
- **WHEN** an API request is missing required files or has invalid filenames
- **THEN** the API route SHALL return an HTTP error before calling the workflow service

#### Scenario: HTTP response formatting after service call
- **WHEN** the workflow service returns a successful API result
- **THEN** the API route SHALL format that result as either a file attachment or API-specific JSON response

#### Scenario: HTTP formatting of service usage error
- **WHEN** the workflow service raises `WorkflowError(code="usage_error")`
- **THEN** the API route SHALL map it to HTTP 400

#### Scenario: HTTP formatting of service processing error
- **WHEN** the workflow service raises `WorkflowError(code="processing_error")`
- **THEN** the API route SHALL map it to HTTP 500
