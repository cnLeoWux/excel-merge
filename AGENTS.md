
# PROJECT KNOWLEDGE BASE

**Generated:** 2026-03-30
**Commit:** 3db1e5b
**Branch:** feature/excel-merger-new-feature

## OVERVIEW
Excel Merge Tool — matches order Excel/CSV files with payment/refund files to populate "支付手续费" column. Supports 3 entry modes (interactive, CLI, Flask API). Also generates sales report period markings and monthly reports. Core stack: Python 3.7+, pandas, openpyxl, xlrd, Flask.

## STRUCTURE
```
./
├── utils.py                # Core business logic (~930 lines): matching, reading, writing, reporting
├── cli.py                  # CLI entry (argparse): order_file payment_file [--month YYYYMM]
├── excel_merge.py          # Interactive entry: file picker from ExcelForHandel/
├── excel_merge_api.py      # Flask API: /merge, /merge/json, /download/<file>, /health
├── setup.py                # Package config: console_scripts excel-merge & excel-merge-cli
├── requirements.txt        # Runtime deps: pandas, openpyxl, xlrd, flask, werkzeug
├── ExcelForHandel/         # Input data directory (26 sample/test files)
├── documents/              # ARCHITECTURE.md, TECHNICAL_DOCS.md, USAGE_EXAMPLES.md
├── openspec/               # OpenSpec config + 7 capability specs (cli-input, cli-output, core-matching, file-io, sales-report, http-api, agent-documentation)
├── dist/                   # Distribution copy of root scripts (not a build artifact)
├── test_*.py               # 4 ad-hoc test scripts in root (no assertions, not real pytest)
├── check_csv.py            # CSV debug helper
├── debug_csv.py            # CSV/encoding debug script
├── create_sample_data.py   # Generates sample files in ExcelForHandel/
├── verify_result.py        # Manual result verification script
└── verify_original.py      # Manual original file verification script
```

## BUILD / LINT / TEST COMMANDS

```bash
# Install
pip install -r requirements.txt
pip install -e .                    # Editable install (enables console_scripts)

# Run application
python excel_merge.py                                           # Interactive mode
python cli.py order.xlsx payment.xlsx                           # CLI basic match (in-place)
python cli.py order.xlsx payment.xlsx --month 202602            # Sales report workflow (in-place)
python excel_merge_api.py                                       # Flask API on 0.0.0.0:5000

# Console scripts (after pip install -e .)
excel-merge                         # Interactive mode
excel-merge-cli                     # CLI mode

# Tests
pip install -r requirements-dev.txt   # Install pytest + pytest-flask
python -m pytest                      # Run full test suite (unit + integration)
python -m pytest tests/unit -v        # Unit tests only (utils.py logic)
python -m pytest tests/integration -v # Integration tests (CLI subprocess + Flask test client)
python -m pytest -k "sales_report"    # Filter by keyword

# No linter/formatter/type-checker configured
```

## CLI USAGE REFERENCE

> **In-place contract**: All merge and sales-report results are written **in place** to the original order file. The CLI does not produce any separate result file or report file. Back up the order file before invoking if you need a copy.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `order_file` | str | *(required)* | Path to the order data file (.xlsx, .xls, .csv) |
| `payment_file` | str | *(required)* | Path to the payment/refund data file (.xlsx, .xls, .csv) |
| `--month` | str | `None` | Target month in `YYYYMM` format (e.g., `202602`); triggers the sales report workflow |
| `--json` | flag | `False` | Output result as JSON envelope to stdout |
| `--quiet` | flag | `False` | Suppress progress logs; only warnings and errors go to stderr |
| `-v`, `--verbose` | count | `0` | Increase verbosity: `-v` = INFO, `-vv` = DEBUG |

### Basic Matching Workflow

```bash
# Match payment fees and write back to order.xlsx in place
python cli.py order.xlsx payment.xlsx

# Console script (after pip install -e .)
excel-merge-cli order.xlsx payment.xlsx
```

Supported file formats: `.xlsx`, `.xls`, `.csv`. Encoding is auto-detected (gbk → utf-8 → gb2312 → latin-1 → utf-8-sig).

### Sales Report Workflow (`--month`)

Triggered by `--month YYYYMM`. Two-phase processing, all writes go to the order file:

1. **Phase 1 — Match & Mark**: Run payment fee matching, then mark the 销售报表账期 column:
   - "全退": duplicate order numbers whose amounts sum to zero
   - "已取消": order status contains "取消" and amount is 0
2. **Phase 2 — Filter & Mark in place**: Compute, in memory, the rows whose 出行日期 falls in a 1-year window of the target month and that are still unmarked, then back-fill `销售报表YYYYMM` into the 销售报表账期 column for those rows. **No `report_YYYYMM.xlsx` file is produced.**

```bash
# Sales report workflow (writes back to order.xlsx in place)
python cli.py order.xlsx payment.xlsx --month 202602
```

Output: `order.xlsx` updated in place with both 支付手续费 and 销售报表账期 columns populated. No separate report file.

### JSON Output Format

Use `--json` to get structured output. The envelope always has three top-level fields: `ok`, `data`, `error`. The shape of `data` is the same regardless of whether `--month` is passed.

**Success** (exit code 0):
```json
{
  "ok": true,
  "data": {
    "output_file": "order.xlsx",
    "statistics": {
      "total_rows": 100,
      "matched_rows": 85,
      "match_rate": "85.00%"
    }
  },
  "error": null
}
```

`output_file` is always equal to the order file path. The `data` object never contains `report_file`, `report_rows`, or `warnings`.

**Failure** (non-zero exit code):
```json
{
  "ok": false,
  "data": null,
  "error": {
    "code": "file_not_found",
    "message": "File 'order.xlsx' does not exist."
  }
}
```

Possible `error.code` values: `file_not_found`, `processing_error`, `unknown_error`.

### Exit Codes

| Code | Constant | Meaning | Typical Trigger |
|------|----------|---------|-----------------|
| 0 | `EXIT_SUCCESS` | Success | Processing completed normally |
| 1 | `EXIT_GENERAL_ERROR` | General Error | Unexpected/unhandled exception |
| 2 | `EXIT_USAGE_ERROR` | Usage Error | Invalid or missing arguments (argparse). Includes passing the **removed** flags `-o`/`--output` or `--output-dir`. |
| 3 | `EXIT_FILE_NOT_FOUND` | File Not Found | Input file does not exist |
| 4 | `EXIT_PROCESSING_ERROR` | Processing Error | Error during matching, parsing, or in-place write of the order file |

### Agent Recommended Usage

For AI Agents and automation scripts, use `--json --quiet` for clean machine-readable output:

```bash
python cli.py order.xlsx payment.xlsx --json --quiet
```

**stdout/stderr separation**:
- **stdout**: JSON envelope (with `--json`) or a single "订单文件已就地更新" summary line (text mode)
- **stderr**: all logs, progress messages, warnings, and errors

Python subprocess integration:
```python
import subprocess, json

result = subprocess.run(
    ["python", "cli.py", "order.xlsx", "payment.xlsx", "--json", "--quiet"],
    capture_output=True, text=True
)

if result.returncode == 0:
    data = json.loads(result.stdout)["data"]
    print(f"Matched {data['statistics']['matched_rows']} rows")
    print(f"Match rate: {data['statistics']['match_rate']}")
elif result.returncode == 3:
    err = json.loads(result.stdout)["error"]
    print(f"File not found: {err['message']}")
else:
    print(f"Failed with exit code {result.returncode}")
```

### Common Error Scenarios

| Scenario | Exit Code | `error.code` | Resolution |
|----------|-----------|--------------|------------|
| Input file does not exist | 3 | `file_not_found` | Verify file path; check cwd or `ExcelForHandel/` |
| Malformed or unreadable file | 4 | `processing_error` | Confirm file is valid .xlsx/.xls/.csv; re-save as UTF-8 if CSV |
| Missing required columns | 4 | `processing_error` | Ensure order file has 订单号, 外部订单号, 订单金额 columns |
| Order file cannot be overwritten (locked / read-only) | 4 | `processing_error` | Close the file in Excel; check filesystem permissions |
| Passing removed flag (`-o`, `--output`, `--output-dir`) | 2 | *(argparse on stderr)* | Remove the flag; back up the order file beforehand if you wanted "save as" |
| Other invalid CLI arguments | 2 | *(argparse on stderr)* | Run `python cli.py --help` to check syntax |

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Matching algorithm | utils.py `process_excel_files()` L189-518 | 20-char exact → P-number → hyphen fallback |
| P-number extraction | utils.py `extract_p_number()` L15 | Regex `r"P\d+"`, case-sensitive |
| File reading (CSV/Excel) | utils.py `read_file_with_appropriate_method()` L39-186 | Encoding fallback chain, comment skipping |
| File writing | utils.py `write_result_file()` L539 | Preserves CSV vs Excel format |
| Sales report period | utils.py `add_sales_report_period()` L572 | Marks 全退 and 已取消 |
| Monthly report filtering | utils.py `filter_unmarked_and_generate_report()` L743 | Filters by 出行日期 window; in-memory only, no file output |
| Full sales workflow | utils.py `process_sales_report_workflow()` L887 | process → mark → filter (in place; no report file) |
| CLI flags | cli.py `main_cli()` | `order_file`, `payment_file`, `--month`, `--json`, `--quiet`, `-v/-vv` |
| Interactive file picker | excel_merge.py `main()` | Lists ExcelForHandel/ contents |
| API endpoints | excel_merge_api.py | POST /merge, POST /merge/json, GET /download/\<f\> |
| Package entry points | setup.py `entry_points` | excel-merge → excel_merge:main, excel-merge-cli → cli:main_cli |
| Architecture docs | documents/ARCHITECTURE.md | System design overview |
| Technical docs | documents/TECHNICAL_DOCS.md | Implementation details |
| Usage examples | documents/USAGE_EXAMPLES.md | Practical examples |
| OpenSpec config | openspec/config.yaml, openspec/project.md | Schema: spec-driven |
| OpenSpec capability specs | openspec/specs/{cli-input,cli-output,core-matching,file-io,sales-report,http-api,agent-documentation}/spec.md | Source of truth for behavior contracts. Run `openspec validate --all --strict` after changes. |

## CODE MAP

### Core Functions (utils.py)

| Symbol | Line | Purpose |
|--------|------|---------|
| `extract_p_number(text)` | 15 | Regex `r"P\d+"` from any input → `Optional[str]` |
| `match_orders_by_p_number(ext_no, prod_name)` | 27 | Compare P-numbers from both fields → `bool` |
| `read_file_with_appropriate_method(file_path)` | 39 | CSV/Excel reader with encoding fallback → `DataFrame` |
| `process_excel_files(order, payment, verbose)` | 189 | **Main matching loop**: exact→P-number→hyphen → `DataFrame` |
| `find_file_path(filename)` | 520 | Search cwd then ExcelForHandel/ → `Path` |
| `write_result_file(df, file_path)` | 539 | Write preserving CSV/Excel format |
| `add_sales_report_period(order_df, verbose)` | 572 | Mark 全退/已取消 in 销售报表账期 column |
| `parse_date(date_val)` | 687 | Multi-format date parser → `Optional[pd.Timestamp]` |
| `get_year_month(date_val)` | 726 | Date → "YYYYMM" string |
| `filter_unmarked_and_generate_report(...)` | 743 | Phase 2: filter unmarked rows, return DataFrames in memory (no file write) |
| `process_sales_report_workflow(...)` | 887 | End-to-end: process + filter + report |

### Dependency Graph
```
cli.py ──────────┐
excel_merge.py ──┤──→ utils.py (all core logic)
excel_merge_api.py┘
                  ↓
            pandas, openpyxl, xlrd, flask
```

### Entry Points
| Script | Function | Console Script |
|--------|----------|----------------|
| excel_merge.py | `main()` | `excel-merge` |
| cli.py | `main_cli()` | `excel-merge-cli` |
| excel_merge_api.py | Flask app | N/A (run directly) |

## CONVENTIONS

### Matching Algorithm (priority order)
1. **Exact match**: first 20 chars of `订单号` ↔ `商户订单号` (column found by substring "商户"+"订单")
2. **P-number match**: `r"P\d+"` extracted from `外部订单号` ↔ `商品名称`
3. **Hyphen match**: `外部订单号` parts ↔ last segment after "-" in `商品名称`
4. **Business type gate**: all matches require type agreement:
   - Regular (订单金额 > 0): payment must be "收费" or "服务费"
   - Refund (订单金额 < 0): payment must be "退费" or "退款"
5. **Amount assignment**:
   - Regular → `支出金额（-元）` (expected negative)
   - Refund → `收入金额（+元）` (expected positive)
   - Zero amount → `支付手续费 = 0.0`, skip matching

### Encoding Fallback Chain (CSV)
`gbk → utf-8 → gb2312 → latin-1 → utf-8-sig`
Then retry with separators `,`, `;`, `\t`, then `sep=None` auto-detect.

### Excel Engine Detection
- `.xlsx`: zipfile check → openpyxl (success) or xlrd (BadZipFile)
- `.xls`: always xlrd

### Column Name Detection
- Business order column: first column where `"商户" in col and "订单" in col`, fallback to any containing `"订单"`
- Order number columns forced to `str` via `astype(str)` or `dtype={"订单号": str}` to prevent numeric conversion
- CSV columns containing `"订单"` or `"流水"` → cast to str after read

### File Handling
- **In-place modification**: the CLI always overwrites the original order file. The `-o`/`--output` and `--output-dir` flags have been removed; copy the file manually if you need a "save as".
- **CSV comments**: lines starting with `#` skipped; first non-comment line = header
- **CSV write**: `utf-8-sig` encoding
- **File discovery**: searches cwd → `ExcelForHandel/` subdirectory

### Code Style
- **Imports**: stdlib → third-party → local (absolute imports only)
- **Naming**: `snake_case` functions/variables, Chinese column names for business fields
- **Indentation**: 4 spaces
- **Type hints**: used on function signatures (`Optional[str]`, `pd.DataFrame`, `Path`)
- **Error handling**: specific exceptions preferred, `verbose: bool = False` pattern for debug output
- **DataFrame**: `pd.isna()`/`pd.notna()` for NA checks, force string types on order columns

### Dependencies
```
pandas>=1.3.0      # on_bad_lines param requires >=1.3
openpyxl>=3.0.0
xlrd>=2.0.0
flask>=2.0.0
werkzeug>=2.0.0
```

## ANTI-PATTERNS (THIS PROJECT)

### Structural
- Flat root layout — no `src/` or package directory, no `__init__.py`
- Tests in root (not `tests/`), ad-hoc scripts with no assertions — NOT real automated tests
- Diagnostic scripts (`check_csv.py`, `debug_csv.py`, `verify_*.py`) mixed with production code
- No CI/CD (.github/workflows), no linting/formatting config, no pyproject.toml, no pytest.ini
- `dist/excel-merge-tool/` is a manual copy of root scripts, not a proper build artifact

### Code Quality
- **Bare/broad exception handlers** throughout utils.py (lines 104, 140, 161, 562, 706, 713), check_csv.py, debug_csv.py — catches `SystemExit`/`KeyboardInterrupt`, hides bugs
- **`on_bad_lines="skip"`** silently drops malformed CSV rows
- **`readlines()` to count comment lines** (utils.py L57-58) — reads entire file into memory
- **Magic number `[:20]`** for order number truncation — should be a named constant
- **`df.shape[1] > 5`** as read-success heuristic — arbitrary threshold
- **`astype(str)` on columns** can convert NaN to literal string `"nan"`
- **print-based logging** instead of `logging` module (utils.py imports logging but uses print)
- **In-place file overwrite by default** — risky for production data

### API-specific
- `/merge` returns fixed XLSX mimetype regardless of actual file format
- Upload/result dirs created at import time in cwd (`uploads/`, `results/`)
- `MAX_CONTENT_LENGTH` declared but may not be enforced via Flask config

## UNIQUE STYLES
- **Multiple engine detection**: zipfile probe to choose openpyxl vs xlrd
- **Column name flexibility**: substring search for Chinese column names (e.g., `"商户" in col and "订单" in col`)
- **Verbose logging**: optional `verbose` flag prints detailed matching progress step-by-step
- **Encoding fallback chain**: tries 5 encodings × 3 separators before giving up
- **Sales report workflow**: two-phase processing (match payments → generate monthly report with 1-year travel date window)
- **P-number regex**: simple `r"P\d+"` — case-sensitive, no separators

## NOTES
- utils.py is the single monolithic module (~930 lines) containing ALL business logic
- `process_excel_files` iterates with `iterrows()` + nested loops — O(n×m) per order×payment, slow for large datasets
- P-number regex `r"P\d+"` is case-sensitive — won't match lowercase `p`
- The 20-char truncation for exact matching is a hardcoded business rule from payment provider format
- Amount columns use full-width parentheses: `支出金额（-元）`, `收入金额（+元）` — exact strings required
- Flask API in debug mode by default (`excel_merge_api.py`)
- `add_sales_report_period` marks "全退" when duplicate order numbers sum to zero, "已取消" when status contains "取消" and amount is 0
