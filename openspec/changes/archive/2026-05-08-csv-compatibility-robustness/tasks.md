## 1. CSV Reading Optimizations

- [x] 1.1 Update `read_file_with_appropriate_method` in `utils.py` to change all instances of `on_bad_lines="skip"` to `on_bad_lines="warn"`.
- [x] 1.2 Update the same function to relax the valid DataFrame heuristic from `df.shape[1] > 5` to `df.shape[1] >= 2`.
- [x] 1.3 Inject `dtype=str` into all `pd.read_csv` calls across `utils.py` to prevent float coercion of large numerical strings.

## 2. String Cleanup Implementation

- [x] 2.1 Update the fallback loops in `read_file_with_appropriate_method` to append `.str.strip('="\t ')` when coercing columns like "订单号", "商户订单号", and "商务订单号" to strings.
- [x] 2.2 Identify any remaining manual casts for order columns in `process_excel_files` or `add_sales_report_period` and ensure they also strip artifact characters if necessary.

## 3. Testing and Verification

- [x] 3.1 Create or modify tests to verify a CSV containing edge cases (e.g. `="12345..."` order numbers and malformed rows) parses successfully without data corruption.
- [x] 3.2 Run the full test suite (`python3 -m pytest tests/`) to ensure the modifications do not break existing CSV or Excel imports.