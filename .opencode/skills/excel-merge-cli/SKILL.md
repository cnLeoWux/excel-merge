---
name: excel-merge-cli
description: Run the excel-merge CLI to match order files with payment/refund files; optionally trigger the sales-report workflow that marks the 销售报表账期 column. All results are written in place to the order file - the CLI never produces any separate output or report file.
license: MIT
metadata:
  author: excel-merge
  version: "2.2"
---

# Excel Merge CLI Skill

Use this skill when the user wants to invoke the project's CLI (`cli.py` / `excel-merge-cli`) to:

1. Match an **order file** with a **payment/refund file** and fill in the `支付手续费` column.
2. Optionally run the **sales report workflow** (`--month YYYYMM`) which marks 全退/已取消 rows and back-fills `销售报表YYYYMM` into the 销售报表账期 column for rows whose 出行日期 falls in a 1-year window of the target month.
3. Produce **machine-readable JSON output** for automation/agent integration.

> **In-place contract**: Every successful invocation writes back to the original order file. The CLI does not produce any separate result file or `report_YYYYMM.xlsx`. If the user wants "save as" behavior, they must copy the order file *before* running the CLI.

This skill is for *invoking* the CLI, not for modifying its source. For source changes, follow the OpenSpec workflow on the `cli-input` / `cli-output` / `sales-report` capabilities.

> **Implementation note**: this skill reflects `cli.py` as of the `remove-output-file-option` change (archived 2026-04-29). If `cli.py` ever re-introduces output flags or report files, update this skill before using it.

---

## When to Use

Trigger this skill when the user says things like:
- "用 CLI 合并这两个文件"
- "Run excel-merge on order.xlsx and payment.xlsx"
- "标注 202602 的销售报表账期" / "Mark the 销售报表账期 column for 202602"
- "Match payment fees from the command line"
- "Give me JSON output from excel-merge"

> **Note on `--month`**: this flag triggers the *sales-report workflow*, which **only marks the 销售报表账期 column inside the order file**. It does **not** generate any `report_YYYYMM.xlsx` or other separate report file. If the user asks for a "sales report file" or "月报文件", clarify the contract before running — the only artifact is the in-place updated order file.

Do **not** use this skill for:
- Interactive mode (`python excel_merge.py`) — that's a TTY file picker.
- HTTP API usage (`excel_merge_api.py`) — different entry point that *does* still produce a downloadable monthly report internally.
- Editing the matching algorithm — that's a code change, not a CLI invocation.

---

## Prerequisites Checklist

Before running, verify:

1. **Working directory** is the project root (where `cli.py` lives).
2. **Dependencies installed**: `pip install -r requirements.txt`. For the `excel-merge-cli` console script, also `pip install -e .`.
3. **Input files exist at the exact path you pass**. The CLI checks `Path(args.order_file).exists()` and `Path(args.payment_file).exists()` directly — it does **not** auto-search `ExcelForHandel/` for bare filenames. If the file lives in `ExcelForHandel/`, pass the full relative path: `python cli.py ExcelForHandel/order.xlsx ExcelForHandel/payment.xlsx`. Use `ls ExcelForHandel/` to confirm.
4. **File formats** are `.xlsx`, `.xls`, or `.csv`. CSV encoding is auto-detected (gbk → utf-8 → gb2312 → latin-1 → utf-8-sig).
5. **Required columns** exist in the order file: `订单号`, `外部订单号`, `订单金额` (and `商品名称` for P-number/hyphen matching). The payment file needs a `商户`+`订单` column, an amount column (`支出金额（-元）` or `收入金额（+元）`), and a business-type column.
6. **Backup, if needed**: Because the CLI overwrites the order file in place, if the user has not made a copy and might want one, recommend `cp order.xlsx order.bak.xlsx` *before* invocation.

If a prerequisite is missing or ambiguous, ask the user before running.

---

## Argument Reference

| Argument | Required | Default | Purpose |
|---|---|---|---|
| `order_file` | yes | — | Order data file (positional #1). Will be overwritten in place. |
| `payment_file` | yes | — | Payment/refund file (positional #2). |
| `--month YYYYMM` | no | `None` | Trigger sales report workflow (e.g. `202602`). Marks 销售报表账期 column in the order file; produces no separate report file. |
| `--json` | no | `False` | Emit JSON envelope on stdout. |
| `--quiet` | no | `False` | Sets the logger to WARNING (suppresses INFO progress); warnings & errors still go to stderr. Does **not** suppress the final stdout summary line. |
| `-v` / `-vv` | no | INFO (default level) | `-v` keeps INFO; `-vv` enables DEBUG. Logging stream is **stderr** (`logging.basicConfig(stream=sys.stderr)`). |

> **Removed flags**: `-o`/`--output` and `--output-dir` no longer exist. Passing them causes argparse to exit with code 2 and an "unrecognized arguments" message on stderr.

**stdout vs stderr**:
- **stderr**: `logging` output (anything routed through the logger), argparse usage errors, the optional traceback in text-mode failures.
- **stdout**: in `--json` mode, exactly the JSON envelope. In text mode, `cli.py` uses plain `print()` for a small set of progress lines (`Processing files:`, `Order file:`, `Payment/Refund file:`, and — when `--month` is used — `执行销售报表工作流...` / `目标月份:`), plus the final "订单文件已就地更新: <path>" summary. The progress lines are suppressed by `--quiet`. The final summary (text mode) or JSON envelope (`--json` mode) is always printed on success regardless of `--quiet`, since `output_result` runs unconditionally.

> ⚠ **Caveat for piping**: in text mode without `--json`, stdout is **not** clean machine-readable output — it interleaves progress lines with the summary. For automation, always use `--json --quiet` so stdout becomes a single JSON object.

---

## Decision Tree

```
Does the user want a monthly sales report?
├── yes → add --month YYYYMM
└── no  → just positional args

Is this for an automation/agent/script?
├── yes → add --json --quiet
└── no  → leave defaults (human-readable text)

Will the original order file be overwritten?
├── always (this is the only mode)
└── if user wants safety → tell them to copy the file BEFORE invoking
```

**Safety nudge**: The CLI **always overwrites the original order file**. There is no opt-out. If you suspect the user might regret this, recommend they back up `order.xlsx` before you run the command.

---

## Canonical Invocations

### 1. Basic match (in-place, the only mode)
```bash
python cli.py order.xlsx payment.xlsx
```
Result: `order.xlsx` updated with 支付手续费 column.

### 2. Sales report workflow
```bash
python cli.py order.xlsx payment.xlsx --month 202602
```
Result: `order.xlsx` updated with both 支付手续费 and 销售报表账期 columns. **No** `report_*.xlsx` is produced anywhere.

### 3. Agent/automation mode (JSON, no log noise)
```bash
python cli.py order.xlsx payment.xlsx --json --quiet
```

### 4. Agent mode with sales report
```bash
python cli.py order.xlsx payment.xlsx --month 202602 --json --quiet
```

### 5. Console-script form (after `pip install -e .`)
```bash
excel-merge-cli order.xlsx payment.xlsx
```

### 6. Safe "save as" pattern (manual, since `-o` is gone)
```bash
cp order.xlsx order_result.xlsx
python cli.py order_result.xlsx payment.xlsx
```

---

## Exit Codes & Error Handling

| Code | Meaning | What to do |
|---|---|---|
| 0 | Success | Parse stdout if `--json`, else read text |
| 1 | General/unknown error | Re-run with `-vv` to capture traceback on stderr |
| 2 | Usage error (argparse) | Run `python cli.py --help`. **Includes** passing the removed `-o`/`--output`/`--output-dir` |
| 3 | File not found | Verify path; check `ExcelForHandel/`; fix typo |
| 4 | Processing error | Confirm columns/format; CSV encoding; **also** raised when the order file cannot be overwritten (locked/read-only) |

When invoked from automation, **always check the exit code first**, then parse JSON.

---

## JSON Output Shape

Always three top-level keys: `ok`, `data`, `error` (one of `data` / `error` is non-null). The shape of `data` is **identical regardless of whether `--month` is passed**.

**Success:**
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

`output_file` is always equal to the order file path. The `data` object **never** contains `report_file`, `report_rows`, or `warnings`.

**Failure:**
```json
{
  "ok": false,
  "data": null,
  "error": { "code": "file_not_found", "message": "File 'x.xlsx' does not exist." }
}
```

Possible `error.code` values emitted by `cli.py` today: `file_not_found`, `processing_error`. (`output_result` also has a defensive `unknown_error` fallback for callers that omit `code`, but no current code path triggers it.)

---

## Recommended Workflow for the Agent

1. **Confirm intent**: basic match vs. sales report. Confirm the user is OK with the order file being overwritten (or recommend a manual copy first).
2. **Locate inputs**: the CLI does **not** auto-search `ExcelForHandel/`. If the user gave bare filenames, check both cwd and `ExcelForHandel/` with Glob/Bash `ls`, then pass the resolved relative path (e.g. `ExcelForHandel/order.xlsx`) to the CLI.
3. **Build the command** following the decision tree.
4. **Run via Bash tool**, capturing stdout/stderr.
5. **Inspect**:
   - exit code first
   - if `--json`: parse stdout, report `match_rate` and `matched_rows` from `data.statistics`
   - else: surface the "订单文件已就地更新" line printed to stdout
6. **On non-zero exit**: re-run with `-vv` (without `--quiet`) to capture detailed stderr, then report the error.
7. **Always remind the user** that the order file was modified in place.

---

## Common Pitfalls

- **20-char truncation**: exact match compares the first 20 chars of `订单号` with `商户订单号`. Truncated/short order numbers may fail to match — that's expected behavior, not a bug.
- **P-number regex is case-sensitive** (`r"P\d+"`). Lowercase `p` won't match.
- **Amount columns use full-width parens**: `支出金额（-元）` and `收入金额（+元）`. Don't substitute half-width `()`.
- **Business-type gating**: regular orders (`订单金额 > 0`) only match `收费`/`服务费`; refunds (`< 0`) only match `退费`/`退款`. Mismatches are skipped silently.
- **CSV with `#` comments**: lines starting with `#` are skipped; first non-comment line is the header.
- **Order file locked in Excel**: writing back will fail with exit code 4 / `processing_error`. Tell the user to close the file in Excel before re-running.
- **No safety net for in-place writes**: passing `-o` / `--output` / `--output-dir` exits with code 2 (argparse "unrecognized arguments") — they must copy the file manually for "save as".
- **Text-mode stdout is not pure**: progress lines and the final summary share stdout (see Argument Reference). Use `--json --quiet` whenever stdout will be parsed.
- **`--quiet` does not silence the summary**: the final "订单文件已就地更新: <path>" line is printed via `print()`, not the logger, so it always appears on success. Redirect stdout if you need true silence.

---

## Quick Reference Card

```bash
# Help
python cli.py --help

# Basic match (in-place)
python cli.py ORDER PAYMENT

# Sales report (in-place; no report file produced)
python cli.py ORDER PAYMENT --month YYYYMM

# Agent/automation
python cli.py ORDER PAYMENT --json --quiet

# Save-as workaround
cp ORDER ORDER_COPY && python cli.py ORDER_COPY PAYMENT

# Debug a failure
python cli.py ORDER PAYMENT -vv
```
