# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Merge Tool matches order Excel/CSV files with payment/refund files to populate the "支付手续费" (Payment Processing Fee) column. It supports sales report period marking and monthly report generation via a two-phase `--month YYYYMM` workflow.

## Commands

```bash
# Install
pip install -r requirements.txt
pip install -e .                    # Editable install (enables console_scripts)

# Run
python excel_merge.py               # Interactive mode (picks files from ExcelForHandel/)
python cli.py order.xlsx payment.xlsx  # CLI - writes back to order file in place
python cli.py order.xlsx payment.xlsx --month 202602  # Sales report workflow (in place)
python excel_merge_api.py           # Flask API on 0.0.0.0:5000

# Console scripts (after pip install -e .)
excel-merge                         # Interactive
excel-merge-cli                     # CLI

# Test
pip install -r requirements-dev.txt
python -m pytest                      # Full suite
python -m pytest tests/unit -v         # Unit tests only
python -m pytest tests/integration -v  # Integration tests only
python -m pytest -k "sales_report"     # Filter by keyword
```

## Architecture

### Feishu Integration

The tool now supports a Feishu workflow via an OpenCode skill:
- **`feishu_workflow.py`** (via `.opencode/skills/excel-merge-cli/SKILL.md`): Wraps the CLI to download uploaded files in Feishu chats, process them in-place, and upload the processed result back to the chat.

### Entry Points

All entry points delegate to `utils.py` which contains all business logic (~930 lines):

```
excel_merge.py ──┐
cli.py ──────────┤──→ utils.py
excel_merge_api.py┘
```

| Entry | Function | Console Script |
|-------|----------|----------------|
| excel_merge.py | `main()` | excel-merge |
| cli.py | `main_cli()` | excel-merge-cli |
| excel_merge_api.py | Flask app | - |

### Core Matching Algorithm (3-tier priority)

1. **Exact match**: First 20 chars of `订单号` ↔ `商户订单号`
2. **P-number match**: `r"P\d+"` extracted from `外部订单号` ↔ `商品名称`
3. **Hyphen match**: `外部订单号` ↔ last segment after `-` in `商品名称`

All matches require business type validation:
- Regular orders (金额 > 0): payment must be "收费" or "服务费" → `支出金额（-元）`
- Refund orders (金额 < 0): payment must be "退费" or "退款" → `收入金额（+元）`

### Sales Report Workflow (`--month YYYYMM`)

Two-phase processing, all writes go to the order file **in place** (no separate report file):

1. **Phase 1 — Match & Mark**: Populate 支付手续费, then mark 销售报表账期 as "全退" (duplicate orders summing to zero) or "已取消" (cancelled with zero amount)
2. **Phase 2 — Filter & Backfill**: Rows whose 出行日期 falls in a 1-year window of target month get `销售报表YYYYMM` written to 销售报表账期

### Key utils.py Functions

| Function | Line | Purpose |
|----------|------|---------|
| `extract_p_number()` | ~18 | Regex `r"P\d+"` extraction |
| `process_excel_files()` | ~192 | Main matching loop (exact→P-number→hyphen) |
| `add_sales_report_period()` | ~582 | Mark 全退/已取消 |
| `filter_unmarked_and_generate_report()` | ~753 | Phase 2 filtering (in-memory, no file output) |
| `process_sales_report_workflow()` | ~883 | End-to-end sales report workflow |

### File Format Handling

- **CSV encoding fallback**: `gbk → utf-8 → gb2312 → latin-1 → utf-8-sig`
- **Excel engine detection**: `.xlsx` uses openpyxl (fallback xlrd), `.xls` uses xlrd
- **In-place contract**: CLI always overwrites the original order file. The `-o`/`--output` and `--output-dir` flags have been removed.

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | General Error |
| 2 | Usage Error (includes passing removed flags) |
| 3 | File Not Found |
| 4 | Processing Error |

### JSON Output

Use `--json --quiet` for machine-readable output. Envelope: `{ok, data, error}` where data contains `output_file` and `statistics` (total_rows, matched_rows, match_rate). The `data` shape is identical regardless of `--month`.

## Documentation

- `AGENTS.md` — Project knowledge base for AI tools
- `USAGE.md` — Chinese usage documentation
- `documents/ARCHITECTURE.md` — Architecture overview
- `documents/TECHNICAL_DOCS.md` — Implementation details
- `openspec/specs/` — OpenSpec capability specs (source of truth for behavior contracts)
