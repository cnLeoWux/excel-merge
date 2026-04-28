# Excel Merge Tool

## Overview

Excel Merge Tool matches order Excel/CSV files with payment/refund files to populate the "支付手续费" (Payment Processing Fee) column. It also supports sales report period marking and monthly report generation.

## Features

- **Multi-tier matching**: 20-char exact match → P-number pattern → hyphen fallback
- **Business type validation**: Regular orders (收费/服务费) vs refund orders (退费/退款)
- **File format support**: Excel (.xlsx, .xls) and CSV with automatic encoding detection
- **Encoding fallback**: gbk → utf-8 → gb2312 → latin-1 → utf-8-sig
- **4 entry modes**: Interactive, CLI, Flask API, console scripts
- **Sales report workflow**: Period marking (全退/已取消) and monthly report generation
- **In-place or output**: Modify original file or specify output path

## Requirements

- Python 3.7+

## Dependencies

```
pandas>=1.3.0
openpyxl>=3.0.0
xlrd>=2.0.0
flask>=2.0.0
werkzeug>=2.0.0
```

## Installation

```bash
pip install -r requirements.txt

# Optional: register console commands
pip install -e .
```

## Usage

### Interactive Mode

```bash
python excel_merge.py
# or after pip install -e .:
excel-merge
```

Lists files in `ExcelForHandel/` for interactive selection.

### CLI Mode

```bash
# Basic: modify original file in-place
python cli.py order.xlsx payment.xlsx

# Specify output file
python cli.py order.xlsx payment.xlsx -o result.xlsx

# Sales report workflow
python cli.py order.xlsx payment.xlsx --month 202602 --output-dir ./reports

# JSON output (for AI Agent integration)
python cli.py order.xlsx payment.xlsx --json

# Quiet mode (suppress logs)
python cli.py order.xlsx payment.xlsx --quiet

# Verbose mode (detailed logs)
python cli.py order.xlsx payment.xlsx -v

# Console script
excel-merge-cli order.xlsx payment.xlsx -o result.xlsx
```

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

### AI Agent / Automation Mode

For AI Agents and automation scripts, use `--json --quiet` for clean machine-readable output:

```bash
# Recommended: JSON output + quiet mode
python cli.py order.xlsx payment.xlsx --json --quiet

# Non-interactive with excel_merge.py
python excel_merge.py --non-interactive --order-file order.xlsx --payment-file payment.xlsx --json

# Check exit code
python cli.py order.xlsx payment.xlsx --json --quiet
echo $?  # 0=success, 3=file not found, 4=processing error
```

**stdout/stderr separation**:
- **stdout**: JSON result only (when `--json`) or result file path (text mode)
- **stderr**: all logs, progress messages, warnings, and errors

#### Exit Codes

| Code | Constant | Meaning | Typical Trigger |
|------|----------|---------|-----------------|
| 0 | `EXIT_SUCCESS` | Success | Processing completed normally |
| 1 | `EXIT_GENERAL_ERROR` | General Error | Unexpected/unhandled exception |
| 2 | `EXIT_USAGE_ERROR` | Usage Error | Invalid or missing arguments (argparse) |
| 3 | `EXIT_FILE_NOT_FOUND` | File Not Found | Input file does not exist |
| 4 | `EXIT_PROCESSING_ERROR` | Processing Error | Error during matching or file writing |

#### JSON Output Format

Success response:
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

Error response:
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

When `--month` is used, `data` also includes `"report_file"` (string or null) and `"report_rows"` (int).

Possible `error.code` values: `file_not_found`, `processing_error`, `unknown_error`.

### Flask API

```bash
python excel_merge_api.py
# Runs on http://localhost:5000
```

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Web test page with upload form |
| GET | `/health` | Health check |
| POST | `/merge` | Upload files, returns processed file |
| POST | `/merge/json` | Upload files, returns JSON with download URL |
| GET | `/download/<file>` | Download result file |

```bash
# Direct download
curl -X POST http://localhost:5000/merge \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.xlsx" \
  --output result.xlsx

# JSON mode
curl -X POST http://localhost:5000/merge/json \
  -F "order_file=@orders.xlsx" \
  -F "payment_file=@payments.csv"
```

## Matching Logic

1. **Exact match**: First 20 chars of 订单号 ↔ 商户订单号
2. **P-number match**: Regex `r"P\d+"` extracted from 外部订单号 ↔ 商品名称
3. **Hyphen match**: 外部订单号 ↔ last segment after `-` in 商品名称
4. **Business type gate**: All matches require type agreement (正单→收费, 退单→退费)
5. **Amount assignment**: 正单→支出金额（-元）, 退单→收入金额（+元）, 零金额→0.0

## Sales Report Workflow

Triggered by `--month YYYYMM`:

1. Match payment fees (same as basic mode)
2. Mark 销售报表账期 column: "全退" (duplicate orders summing to zero), "已取消" (cancelled status with zero amount)
3. Filter unmarked rows with 出行日期 within 1-year window of target month
4. Generate `report_YYYYMM.xlsx`

## Project Structure

```
excel-merge/
├── utils.py                # Core business logic (~930 lines)
├── cli.py                  # CLI entry point (argparse)
├── excel_merge.py          # Interactive entry point
├── excel_merge_api.py      # Flask API server
├── setup.py                # Package config with console_scripts
├── requirements.txt        # Runtime dependencies
├── AGENTS.md               # Project knowledge base (for AI tools)
├── USAGE.md                # Chinese usage documentation
├── documents/
│   ├── ARCHITECTURE.md     # Architecture overview
│   ├── TECHNICAL_DOCS.md   # Implementation details
│   └── USAGE_EXAMPLES.md   # Usage examples
├── openspec/               # OpenSpec configuration
├── ExcelForHandel/         # Input data directory
└── dist/                   # Distribution copy of scripts
```

## Documentation

- [USAGE.md](USAGE.md) — Chinese usage guide (使用文档)
- [documents/ARCHITECTURE.md](documents/ARCHITECTURE.md) — System architecture
- [documents/TECHNICAL_DOCS.md](documents/TECHNICAL_DOCS.md) — Technical implementation details
- [documents/USAGE_EXAMPLES.md](documents/USAGE_EXAMPLES.md) — Detailed usage examples with sample data

## Specifications (OpenSpec)

This project uses [OpenSpec](https://github.com/Fission-AI/OpenSpec) for spec-driven development. All capability contracts (CLI, matching, file I/O, sales report, HTTP API, agent documentation) live in `openspec/specs/`. Use these specs as the source of truth when changing behavior.

- `openspec/project.md` — Project context, conventions, and constraints
- `openspec/specs/cli-input/` — Non-interactive CLI invocation contract
- `openspec/specs/cli-output/` — JSON envelope, exit codes, stdout/stderr separation
- `openspec/specs/core-matching/` — Multi-tier matching algorithm and business type validation
- `openspec/specs/file-io/` — Encoding fallback, Excel engine detection, order-number string protection
- `openspec/specs/sales-report/` — Two-phase sales report workflow (`--month`)
- `openspec/specs/http-api/` — Flask endpoints, upload/download contracts
- `openspec/specs/agent-documentation/` — AGENTS.md content requirements

```bash
# List specs and changes
openspec list --specs
openspec list

# Validate all specs
openspec validate --all --strict

# View a specific spec
openspec show cli-output
```
