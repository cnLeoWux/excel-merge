## 1. API Implementation (`excel_merge_api.py`)

- [x] 1.1 Modify the `/merge` and `/merge/json` endpoints to read an optional `month` parameter from the request form data.
- [x] 1.2 In both endpoints, add a conditional check. If `month` is provided, call `utils.process_sales_report_workflow`. Otherwise, call the existing `utils.process_excel_files`.
- [x] 1.3 For the `/merge/json` endpoint, when `month` is used, ensure the returned JSON `data` object includes the `report_file` and `report_rows` fields, consistent with the CLI.

## 2. Interactive Mode Implementation (`excel_merge.py`)

- [x] 2.1 In the `main` function, after the order and payment files are selected, add a prompt asking the user "Do you want to generate a sales report? (y/n)".
- [x] 2.2 If the user responds 'y', prompt them to enter the month with "Enter the report month (e.g., 202602): ".
- [x] 2.3 Add a loop to validate the user's month input, ensuring it matches the `YYYYMM` format.
- [x] 2.4 If the month is valid, call `utils.process_sales_report_workflow` with the correct parameters.

## 3. Verification

- [x] 3.1 Manually run `python excel_merge.py` and test the new interactive sales report flow (both 'y' and 'n' paths).
- [x] 3.2 Use `curl` to test the `/merge/json` endpoint with and without the `month` parameter, and verify the JSON output is correct in both cases.
- [x] 3.3 Use `curl` to test the `/merge` endpoint with the `month` parameter and verify the correct report file is downloaded.
- [x] 3.4 Run `python cli.py --month 202602 ...` to ensure that the existing CLI functionality is unaffected.

## 4. Documentation

- [x] 4.1 Update `README.md` and/or `documents/USAGE_EXAMPLES.md` to document the new `month` parameter for the API endpoints.
- [x] 4.2 Update the same documents to describe the new interactive sales report generation flow.
