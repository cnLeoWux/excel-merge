## Why

The project currently relies on ad-hoc manual scripts for testing, which lack assertions and a structured framework. This makes it difficult to verify code changes, catch regressions, and ensure the tool's reliability. Introducing an automated test suite is crucial for improving code quality and long-term maintainability.

## What Changes

- Introduce `pytest` as the official testing framework.
- Create a new `tests/` directory to house all test files, separating them from application code.
- Convert the existing `test_*.py` scripts into proper `pytest` tests with assertions.
- Add comprehensive unit and integration tests for the core logic in `utils.py`, covering the matching algorithm, sales report workflow, and file I/O.
- Add functional tests for the CLI (`cli.py`) and interactive (`excel_merge.py`) entry points.
- Add integration tests for the Flask API (`excel_merge_api.py`) endpoints.
- Add `pytest` to a new `requirements-dev.txt` file.
- Update `README.md` to include instructions for installing and running the test suite.

## Capabilities

### New Capabilities
- `automated-testing`: Defines the requirements for the new test suite, including the testing framework, file structure, test coverage, and execution commands.

### Modified Capabilities
- None

## Impact

- **Code**: A new `tests/` directory will be created. Existing `test_*.py` scripts in the root will be moved and refactored.
- **Dependencies**: `pytest` will be added as a development dependency.
- **Documentation**: `README.md` and `AGENTS.md` will be updated to reflect the new testing setup.
- **Development Workflow**: A new step for running `pytest` will be added to the development and verification process.
