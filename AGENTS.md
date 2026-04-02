
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
├── cli.py                  # CLI entry (argparse): order_file payment_file [-o] [--month] [--output-dir]
├── excel_merge.py          # Interactive entry: file picker from ExcelForHandel/
├── excel_merge_api.py      # Flask API: /merge, /merge/json, /download/<file>, /health
├── setup.py                # Package config: console_scripts excel-merge & excel-merge-cli
├── requirements.txt        # Runtime deps: pandas, openpyxl, xlrd, flask, werkzeug
├── ExcelForHandel/         # Input data directory (26 sample/test files)
├── documents/              # ARCHITECTURE.md, TECHNICAL_DOCS.md, USAGE_EXAMPLES.md
├── openspec/               # OpenSpec config (config.yaml, project.md, empty changes/specs)
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
python cli.py order.xlsx payment.xlsx                           # CLI basic
python cli.py order.xlsx payment.xlsx -o result.xlsx            # CLI with output
python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports  # Sales report workflow
python excel_merge_api.py                                       # Flask API on 0.0.0.0:5000

# Console scripts (after pip install -e .)
excel-merge                         # Interactive mode
excel-merge-cli                     # CLI mode

# Tests (ad-hoc scripts, NOT real pytest suites — no assertions)
python -m pytest                    # Collects test_*.py but they just print, no real assertions
python test_engine.py               # Manual: engine detection smoke test
python test_csv_reading.py          # Manual: CSV reading smoke test

# No linter/formatter/type-checker configured
```

## CLI USAGE REFERENCE

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `order_file` | str | *(required)* | Path to the order data file (.xlsx, .xls, .csv) |
| `payment_file` | str | *(required)* | Path to the payment/refund data file (.xlsx, .xls, .csv) |
| `-o`, `--output` | str | `None` (overwrite original) | Output file path; if omitted, the original order file is modified in-place |
| `--month` | str | `None` | Target month in `YYYYMM` format (e.g., `202602`); triggers the sales report workflow |
| `--output-dir` | str | `None` (current dir) | Output directory for the generated monthly report |
| `--json` | flag | `False` | Output result as JSON envelope to stdout |
| `--quiet` | flag | `False` | Suppress progress logs; only warnings and errors go to stderr |
| `-v`, `--verbose` | count | `0` | Increase verbosity: `-v` = INFO, `-vv` = DEBUG |

### Basic Matching Workflow

```bash
# Modify original order file in-place (default)
python cli.py order.xlsx payment.xlsx

# Specify output file
python cli.py order.xlsx payment.xlsx -o result.xlsx

# Console script (after pip install -e .)
excel-merge-cli order.xlsx payment.xlsx -o result.xlsx
```

Supported file formats: `.xlsx`, `.xls`, `.csv`. Encoding is auto-detected (gbk → utf-8 → gb2312 → latin-1 → utf-8-sig).

### Sales Report Workflow (`--month`)

Triggered by `--month YYYYMM`. Two-phase processing:

1. **Phase 1 — Match & Mark**: Run payment fee matching, then mark 销售报表账期 column:
   - "全退": duplicate order numbers whose amounts sum to zero
   - "已取消": order status contains "取消" and amount is 0
2. **Phase 2 — Filter & Generate Report**: Filter unmarked rows with 出行日期 within a 1-year window of the target month, then write `report_YYYYMM.xlsx`.

```bash
# Full sales report workflow
python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports

# Also redirect the updated order file
python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports -o updated_order.xlsx
```

Output: original order file updated (or written to `-o`) + `report_YYYYMM.xlsx` in `--output-dir` (or cwd).

### JSON Output Format

Use `--json` to get structured output. The envelope always has three top-level fields: `ok`, `data`, `error`.

**Success** (exit code 0):
```json
{
  "ok": true,
  "data": {
    "output_file": "result.xlsx",
    "statistics": {
      "total_rows": 100,
      "matched_rows": 85,
      "match_rate": "85.00%"
    }
  },
  "error": null
}
```

When `--month` is used, `data` also includes `"report_file"` (string or null) and `"report_rows"` (int).

**Error** (exit code 3 or 4):
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
| 2 | `EXIT_USAGE_ERROR` | Usage Error | Invalid or missing arguments (argparse) |
| 3 | `EXIT_FILE_NOT_FOUND` | File Not Found | Input file does not exist |
| 4 | `EXIT_PROCESSING_ERROR` | Processing Error | Error during matching or file writing |

### Agent Recommended Usage

For AI Agents and automation scripts, use `--json --quiet` for clean machine-readable output:

```bash
python cli.py order.xlsx payment.xlsx --json --quiet
```

**stdout/stderr separation**:
- **stdout**: JSON result only (when `--json`) or result file path (text mode)
- **stderr**: all logs, progress messages, warnings, and errors

Python subprocess integration:
```python
import subprocess, json

result = subprocess.run(
    ["python", "cli.py", "order.xlsx", "payment.xlsx", "--json", "--quiet"],
    capture_output=True, text=True
)

if result.returncode == 0:
    data = json.loads(result.stdout)
    print(f"Matched {data['data']['statistics']['matched_rows']} rows")
    print(f"Match rate: {data['data']['statistics']['match_rate']}")
elif result.returncode == 3:
    err = json.loads(result.stdout)
    print(f"File not found: {err['error']['message']}")
else:
    print(f"Failed with exit code {result.returncode}")
```

### Common Error Scenarios

| Scenario | Exit Code | `error.code` | Resolution |
|----------|-----------|--------------|------------|
| Input file does not exist | 3 | `file_not_found` | Verify file path; check cwd or `ExcelForHandel/` |
| Malformed or unreadable file | 4 | `processing_error` | Confirm file is valid .xlsx/.xls/.csv; re-save as UTF-8 if CSV |
| Missing required columns | 4 | `processing_error` | Ensure order file has 订单号, 外部订单号, 订单金额 columns |
| Invalid CLI arguments | 2 | *(argparse prints to stderr)* | Run `python cli.py --help` to check syntax |

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Matching algorithm | utils.py `process_excel_files()` L189-518 | 20-char exact → P-number → hyphen fallback |
| P-number extraction | utils.py `extract_p_number()` L15 | Regex `r"P\d+"`, case-sensitive |
| File reading (CSV/Excel) | utils.py `read_file_with_appropriate_method()` L39-186 | Encoding fallback chain, comment skipping |
| File writing | utils.py `write_result_file()` L539 | Preserves CSV vs Excel format |
| Sales report period | utils.py `add_sales_report_period()` L572 | Marks 全退 and 已取消 |
| Monthly report generation | utils.py `filter_unmarked_and_generate_report()` L743 | Filters by 出行日期 window |
| Full sales workflow | utils.py `process_sales_report_workflow()` L887 | process → filter → report |
| CLI flags | cli.py `main_cli()` | `-o`, `--month`, `--output-dir` |
| Interactive file picker | excel_merge.py `main()` | Lists ExcelForHandel/ contents |
| API endpoints | excel_merge_api.py | POST /merge, POST /merge/json, GET /download/\<f\> |
| Package entry points | setup.py `entry_points` | excel-merge → excel_merge:main, excel-merge-cli → cli:main_cli |
| Architecture docs | documents/ARCHITECTURE.md | System design overview |
| Technical docs | documents/TECHNICAL_DOCS.md | Implementation details |
| Usage examples | documents/USAGE_EXAMPLES.md | Practical examples |
| OpenSpec config | openspec/config.yaml, openspec/project.md | Schema: spec-driven |

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
| `filter_unmarked_and_generate_report(...)` | 743 | Phase 2: filter unmarked rows, write report_YYYYMM.xlsx |
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
- **In-place modification**: default behavior overwrites original order file (use `-o` to redirect)
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
