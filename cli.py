import json
import logging
import sys
from pathlib import Path

import pandas as pd

from utils import (
    add_sales_report_period,
    filter_unmarked_and_generate_report,
    find_file_path,
    process_excel_files,
    process_sales_report_workflow,
    read_file_with_appropriate_method,
    write_result_file,
)

# Exit codes
EXIT_SUCCESS = 0
EXIT_GENERAL_ERROR = 1
EXIT_USAGE_ERROR = 2
EXIT_FILE_NOT_FOUND = 3
EXIT_PROCESSING_ERROR = 4


def output_result(data=None, error=None, json_mode=False):
    """
    Output result in JSON or text format.

    JSON mode: outputs envelope (`{ok, data, error}`) to stdout.
    Text mode: prints a short human-readable summary.

    The CLI guarantees that all merge / sales-report results are written
    in place to the original order file. There is no separate "result file"
    or "report file" — `data["output_file"]` always equals the order file path.
    """
    if json_mode:
        if error:
            result = {
                "ok": False,
                "data": None,
                "error": {
                    "code": error.get("code", "unknown_error"),
                    "message": error.get("message", "Unknown error"),
                },
            }
        else:
            result = {
                "ok": True,
                "data": data,
                "error": None,
            }
        print(json.dumps(result, ensure_ascii=False))
    else:
        # Text mode
        if error:
            print(f"错误: {error['message']}")

        if data:
            if "output_file" in data:
                print(f"订单文件已就地更新: {data['output_file']}")

            if "message" in data:
                print(data["message"])


def _build_success_payload(order_file: str, result_df: pd.DataFrame) -> dict:
    """Construct the JSON `data` payload for a successful run.

    Shape is identical for the basic match path and the sales-report path:
    `output_file` always equals the order file (results are in-place),
    plus a `statistics` block.
    """
    total_rows = len(result_df)
    matched_rows = int(result_df["支付手续费"].notna().sum()) if "支付手续费" in result_df.columns else 0
    match_rate = (
        f"{(matched_rows / total_rows * 100):.2f}%" if total_rows > 0 else "0.00%"
    )
    return {
        "output_file": str(order_file),
        "statistics": {
            "total_rows": total_rows,
            "matched_rows": matched_rows,
            "match_rate": match_rate,
        },
    }


def main_cli():
    import argparse

    parser = argparse.ArgumentParser(
        description=(
            "Merge two Excel/CSV files based on payment-fee matching logic. "
            "All results are written in place to the order file; no separate "
            "output or report files are produced."
        )
    )
    parser.add_argument(
        "order_file", type=str, help="Path to the order data file (.xlsx/.xls/.csv)"
    )
    parser.add_argument(
        "payment_file",
        type=str,
        help="Path to the payment/refund data file (.xlsx/.xls/.csv)",
    )

    # 销售报表工作流：仅触发账期标注，不再落盘
    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help="Target month for sales report workflow (format: YYYYMM, e.g., 202602)",
    )

    # JSON output and logging control
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output result as JSON to stdout",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase verbosity (use -v for INFO, -vv for DEBUG)",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress non-error output (only warnings and errors)",
    )

    args = parser.parse_args()

    # Configure logging based on flags
    if args.quiet:
        log_level = logging.WARNING
    elif args.verbose >= 2:
        log_level = logging.DEBUG
    elif args.verbose >= 1:
        log_level = logging.INFO
    else:
        log_level = logging.INFO  # Default level for backward compatibility

    logging.basicConfig(
        level=log_level,
        format="%(message)s",
        stream=sys.stderr,  # Log to stderr
    )

    # Check if files exist
    if not Path(args.order_file).exists():
        output_result(
            error={
                "code": "file_not_found",
                "message": f"File '{args.order_file}' does not exist.",
            },
            json_mode=args.json,
        )
        sys.exit(EXIT_FILE_NOT_FOUND)

    if not Path(args.payment_file).exists():
        output_result(
            error={
                "code": "file_not_found",
                "message": f"File '{args.payment_file}' does not exist.",
            },
            json_mode=args.json,
        )
        sys.exit(EXIT_FILE_NOT_FOUND)

    # Only print file info in non-JSON mode without --quiet
    if not args.json and not args.quiet:
        print(f"Processing files:")
        print(f"  Order file: {args.order_file}")
        print(f"  Payment/Refund file: {args.payment_file}")

    try:
        # Determine verbosity for utils functions
        verbose = args.verbose >= 1 and not args.quiet

        # 销售报表工作流（--month）与基本匹配路径共享同一种产出契约：
        # 所有结果就地写回订单文件；不产生任何独立的结果文件或报表文件。
        if args.month:
            if not args.json and not args.quiet:
                print(f"\n执行销售报表工作流...")
                print(f"  目标月份: {args.month}")

            updated_df, _report_df = process_sales_report_workflow(
                args.order_file,
                args.payment_file,
                args.month,
                verbose=verbose,
            )
            result_df = updated_df
        else:
            # 仅匹配支付手续费
            result_df = process_excel_files(
                args.order_file, args.payment_file, verbose=verbose
            )

        # 统一就地写回订单文件；任何写入异常 → processing_error
        try:
            write_result_file(result_df, Path(args.order_file))
        except Exception as e:
            output_result(
                error={
                    "code": "processing_error",
                    "message": f"无法写回订单文件 '{args.order_file}': {e}",
                },
                json_mode=args.json,
            )
            sys.exit(EXIT_PROCESSING_ERROR)

        output_result(
            data=_build_success_payload(args.order_file, result_df),
            json_mode=args.json,
        )

        sys.exit(EXIT_SUCCESS)

    except SystemExit:
        raise
    except Exception as e:
        output_result(
            error={
                "code": "processing_error",
                "message": str(e),
            },
            json_mode=args.json,
        )
        if not args.json:
            import traceback
            traceback.print_exc()
        sys.exit(EXIT_PROCESSING_ERROR)


if __name__ == "__main__":
    main_cli()
