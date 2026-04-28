## 1. Logic Refactoring (utils.py)

- [x] 1.1 Review `process_sales_report_workflow` to ensure it remains a pure logic orchestrator and doesn't introduce side effects like file saving (currently it doesn't, but good to verify).

## 2. CLI Implementation (cli.py)

- [x] 2.1 Refactor the `--month` branch in `main_cli()` to change the order of operations: generate report first, then attempt to save the order file.
- [x] 2.2 Wrap the `write_result_file` call for the updated order file in a `try-except` block.
- [x] 2.3 Implement warning collection mechanism to track when file saving fails but report generation succeeds.
- [x] 2.4 Update `output_result` or its invocation to include a `warnings` field in the JSON data when partial failures occur.
- [x] 2.5 Ensure text mode output clearly distinguishes between report success and order file failure.

## 3. Verification & Documentation

- [x] 3.1 Verify behavior by simulating a file lock on the order file (e.g., keep it open in Excel) and ensuring the report is still generated.
- [x] 3.2 Run `openspec validate --all --strict` to ensure implementation matches updated specs.
- [x] 3.3 Update `AGENTS.md` or other docs if the exit code or JSON envelope structure changed significantly (specifically adding `warnings`).
