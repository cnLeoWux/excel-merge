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

| Flag | Description |
|------|-------------|
| `order_file` | Path to order data file (required) |
| `payment_file` | Path to payment/refund file (required) |
| `-o`, `--output` | Output file path (default: overwrite original) |
| `--month` | Target month YYYYMM, triggers sales report workflow |
| `--output-dir` | Output directory for generated report |
| `--json` | Output result as JSON to stdout |
| `--quiet` | Suppress progress logs (only errors) |
| `-v`, `--verbose` | Increase verbosity (-v=INFO, -vv=DEBUG) |

### AI Agent / Automation Mode

For AI Agents and automation scripts, use non-interactive mode with JSON output:

```bash
# Non-interactive with JSON output
python cli.py order.xlsx payment.xlsx --json --quiet

# Non-interactive with excel_merge.py
python excel_merge.py --non-interactive --order-file order.xlsx --payment-file payment.xlsx --json

# Check exit code
python cli.py order.xlsx payment.xlsx --json --quiet
echo $?  # 0=success, 3=file not found, 4=processing error
```

#### Exit Codes

| Code | Meaning | Description |
|------|---------|-------------|
| 0 | Success | Processing completed successfully |
| 1 | General Error | Unexpected error occurred |
| 2 | Usage Error | Invalid arguments or parameters |
| 3 | File Not Found | Input file does not exist |
| 4 | Processing Error | Error during matching or writing |

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
