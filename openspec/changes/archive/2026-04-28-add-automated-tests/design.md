## Context

The project currently has no automated test suite. Testing is performed via four ad-hoc scripts in the root directory (`test_engine.py`, `test_csv_reading.py`, etc.) that lack assertions and a formal structure. This makes verifying changes and preventing regressions a manual and error-prone process. The proposal and specs for the `add-automated-tests` change mandate the introduction of `pytest` and the creation of a comprehensive test suite.

## Goals / Non-Goals

**Goals:**
- Establish a robust, automated testing foundation for the project.
- Ensure all core logic, entry points, and API endpoints are covered by tests.
- Make tests easy to run for developers and in future CI/CD pipelines.
- Separate test code from application code.

**Non-Goals:**
- Achieving 100% test coverage. The initial focus is on critical paths and core functionality.
- Implementing a CI/CD pipeline in this change. This design focuses solely on creating the test suite itself.
- Performance or load testing. The tests will focus on correctness and functionality.

## Decisions

1.  **Testing Framework**: `pytest`
    -   **Rationale**: `pytest` is the de-facto standard for testing in the Python ecosystem. It has a rich ecosystem of plugins, requires less boilerplate than the standard `unittest` module, and its fixture model is excellent for managing test setup and teardown (e.g., creating sample files, managing the Flask app context). This aligns with the `automated-testing` spec.
    -   **Alternatives**: `unittest` (too verbose), `nose2` (less common than `pytest`).

2.  **Test Structure**: A new top-level `tests/` directory.
    -   **Rationale**: Separating test code from application code is a standard best practice. It keeps the root directory clean and makes it easy to run all tests by simply targeting the `tests/` directory. This aligns with the `automated-testing` spec. The existing `test_*.py` files in the root will be moved into this new directory and refactored.
    -   **File Naming**: Test files will be named `test_*.py`. This is the standard discovery pattern for `pytest`.

3.  **Test Data Management**: A `tests/conftest.py` file and a `tests/fixtures/` directory.
    -   **Rationale**: `pytest` fixtures are the ideal way to manage test data and shared setup.
    -   **`conftest.py`**: Will contain fixtures for generating sample dataframes, creating temporary files/directories, and providing a test client for the Flask API.
    -   **`tests/fixtures/`**: Will store any static sample files (e.g., pre-defined Excel/CSV files for specific edge cases) that are loaded by the fixtures.

4.  **Test Categories**:
    -   **Unit Tests (`tests/unit/`)**: For testing individual functions in `utils.py` in isolation. These tests will not perform file I/O but will work with in-memory pandas DataFrames.
    -   **Integration Tests (`tests/integration/`)**: For testing the integration between different parts of the system.
        -   `test_cli.py`: Will use `subprocess` to run `cli.py` and assert on exit codes, stdout, and created files.
        -   `test_api.py`: Will use the `pytest-flask` plugin or a custom fixture to get a `test_client` and make HTTP requests to the running Flask application.
    -   **Rationale**: This categorization helps distinguish between fast, isolated unit tests and slower, more complex integration tests.

5.  **Dependencies**: `pytest` and `pytest-flask` will be added to a new `requirements-dev.txt`.
    -   **Rationale**: Separating development dependencies from production dependencies is standard practice. `pytest-flask` simplifies testing Flask applications.

## Risks / Trade-offs

-   **Risk**: The existing ad-hoc test scripts might contain implicit knowledge about edge cases.
    -   **Mitigation**: Carefully review the existing scripts (`test_engine.py`, `test_csv_reading.py`, etc.) during the refactoring process to ensure no test cases are lost. They will be converted into `pytest` tests with explicit assertions.
-   **Risk**: Setting up the API tests might be complex due to file uploads.
    -   **Mitigation**: Use the Flask test client's ability to send `multipart/form-data` requests, simulating a real file upload without needing to run a live server on a network port. The test client operates in-memory.

## Migration Plan

1.  Create the `tests/` directory.
2.  Create `requirements-dev.txt` and add `pytest` and `pytest-flask`.
3.  Create `tests/conftest.py` to define core fixtures.
4.  Move and refactor the existing `test_*.py` scripts into the new `tests/integration/` directory, adding assertions.
5.  Add new unit tests for `utils.py` in `tests/unit/`.
6.  Update `README.md` and `AGENTS.md` with instructions on how to install dev dependencies and run the tests.
7.  Delete the old `test_*.py` scripts from the root directory.
