## Why

Currently, input files can be in `.csv` format (like exported payment logs), but reading these files faces encoding and formatting issues (e.g. varying encodings, scientific notation prefixes like `="123"`). To prevent data loss and ensure robust matching, all input methods (CLI arguments, interactive file picker, HTTP API) need to uniformly and safely handle CSV files, improving overall system resilience.

## What Changes

- Improve CSV parsing logic to handle missing data or formatting edge cases without silent drops.
- Clean up artifacts from payment/order CSVs (e.g., stripping `="` or `\t`).
- Ensure large numbers (like 20-digit order numbers) are correctly parsed as strings instead of falling back to floats which corrupts them.
- Standardize CSV handling across all input entry points (Interactive CLI, Argparse CLI, Flask API).

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `file-io`: Strengthen CSV reading with `dtype=str`, prevent data loss with `on_bad_lines="warn"`, and strip artifacts like `="\t ` across the file reading pipeline.
- `cli-input`: Ensure interactive mode properly handles or supports CSV inputs without edge case failures.
- `http-api`: Ensure file uploads processing properly delegates to the robust `file-io` mechanism when CSVs are uploaded.

## Impact

- `utils.py` (specifically `read_file_with_appropriate_method`) will be fully enhanced to process CSV securely.
- Any interface passing files (`excel_merge.py`, `cli.py`, `excel_merge_api.py`) will implicitly benefit from enhanced robustness. No major breaking changes to APIs, but improved correctness.