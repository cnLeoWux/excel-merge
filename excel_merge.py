"""
Author: Leo Wu leo.wux@lego.com
Date: 2025-10-27 10:59:13
LastEditors: Leo Wu leo.wux@lego.com
LastEditTime: 2025-12-30 13:23:11
FilePath: /excel-merge/excel_merge.py
Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
"""

import argparse
import json
import logging
import sys
from pathlib import Path

import pandas as pd

from utils import (
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
        if error:
            print(f"Error: {error['message']}")
        elif data and "message" in data:
            print(data["message"])


def main():
    # Configure argparse
    parser = argparse.ArgumentParser(
        description="Excel Merge Tool - Interactive or non-interactive mode."
    )

    # Non-interactive mode parameters
    parser.add_argument(
        "--order-file",
        type=str,
        help="Path to the order file (required in non-interactive mode)",
    )
    parser.add_argument(
        "--payment-file",
        type=str,
        help="Path to the payment/refund file (required in non-interactive mode)",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Force non-interactive mode (skip input prompts)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output result as JSON to stdout",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        help="Output file path (default: modify original file)",
    )

    # Logging control
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase verbosity",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress non-error output",
    )

    # Sales report workflow parameters
    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help="Target month for sales report (format: YYYYMM, e.g., 202602)",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Output directory for the generated report",
    )

    args = parser.parse_args()

    # Configure logging
    if args.quiet:
        log_level = logging.WARNING
    elif args.verbose >= 2:
        log_level = logging.DEBUG
    elif args.verbose >= 1:
        log_level = logging.INFO
    else:
        log_level = logging.INFO

    logging.basicConfig(
        level=log_level,
        format="%(message)s",
        stream=sys.stderr,
    )

    # Determine if we are in non-interactive mode
    is_non_interactive = args.non_interactive or not sys.stdin.isatty()

    if is_non_interactive:
        # Non-interactive mode: use command-line arguments
        if not args.order_file or not args.payment_file:
            output_result(
                error={
                    "code": "usage_error",
                    "message": "Non-interactive mode requires --order-file and --payment-file arguments.",
                },
                json_mode=args.json,
            )
            sys.exit(EXIT_USAGE_ERROR)

        order_file_path = Path(args.order_file)
        payment_file_path = Path(args.payment_file)

        # Check if files exist
        if not order_file_path.exists():
            output_result(
                error={
                    "code": "file_not_found",
                    "message": f"Order file '{args.order_file}' does not exist.",
                },
                json_mode=args.json,
            )
            sys.exit(EXIT_FILE_NOT_FOUND)

        if not payment_file_path.exists():
            output_result(
                error={
                    "code": "file_not_found",
                    "message": f"Payment file '{args.payment_file}' does not exist.",
                },
                json_mode=args.json,
            )
            sys.exit(EXIT_FILE_NOT_FOUND)
    else:
        # Interactive mode: use file picker
        print("Excel Merge Tool")

        # Get all files in ExcelForHandel directory
        excel_dir = Path("ExcelForHandel")
        if not excel_dir.exists():
            print(f"Error: ExcelForHandel directory does not exist.")
            sys.exit(EXIT_FILE_NOT_FOUND)

        # List all files in the directory
        files = list(excel_dir.glob("*"))
        if not files:
            print(f"Error: No files found in ExcelForHandel directory.")
            sys.exit(EXIT_FILE_NOT_FOUND)

        # Display files for selection
        print("Available files in ExcelForHandel directory:")
        for i, file in enumerate(files, 1):
            print(f"{i}. {file.name}")

        # Get first file selection
        while True:
            try:
                order_choice = int(
                    input(
                        "\nSelect the first Excel file (order data) by number: "
                    ).strip()
                )
                if 1 <= order_choice <= len(files):
                    order_file_path = files[order_choice - 1]
                    break
                else:
                    print(f"Please enter a number between 1 and {len(files)}.")
            except ValueError:
                print("Please enter a valid number.")

        # Get second file selection
        while True:
            try:
                payment_choice = int(
                    input(
                        "\nSelect the second Excel file (payment/refund data) by number: "
                    ).strip()
                )
                if 1 <= payment_choice <= len(files):
                    payment_file_path = files[payment_choice - 1]
                    break
                else:
                    print(f"Please enter a number between 1 and {len(files)}.")
            except ValueError:
                print("Please enter a valid number.")

    # Processing
    if not args.json and not args.quiet:
        print(f"Processing files:")
        print(f"  Order file: {order_file_path}")
        print(f"  Payment/Refund file: {payment_file_path}")

    try:
        # Determine verbosity
        verbose = args.verbose >= 1 and not args.quiet

        # Check if sales report workflow is requested
        if args.month:
            # Full sales report workflow
            if not args.json and not args.quiet:
                print(f"\n执行完整销售报表工作流...")
                print(f"  目标月份: {args.month}")

            updated_df, report_df = process_sales_report_workflow(
                str(order_file_path),
                str(payment_file_path),
                args.month,
                output_dir=args.output_dir,
                verbose=verbose,
            )

            # Save result
            if args.output:
                output_path = Path(args.output)
                write_result_file(updated_df, output_path)
                result_file = args.output
            else:
                write_result_file(updated_df, order_file_path)
                result_file = str(order_file_path)

            # Report file info
            report_file = None
            if len(report_df) > 0:
                report_filename = f"report_{args.month}.xlsx"
                report_file = str(Path(args.output_dir or ".") / report_filename)
                if not args.json and not args.quiet:
                    print(f"\n新报表文件: {report_filename}")
                    print(f"包含 {len(report_df)} 行数据")

            # Calculate statistics
            total_rows = len(updated_df)
            matched_rows = updated_df["支付手续费"].notna().sum()
            match_rate = (
                f"{(matched_rows / total_rows * 100):.2f}%"
                if total_rows > 0
                else "0.00%"
            )

            if args.json:
                output_result(
                    data={
                        "output_file": result_file,
                        "statistics": {
                            "total_rows": total_rows,
                            "matched_rows": int(matched_rows),
                            "match_rate": match_rate,
                        },
                        "report_file": report_file,
                        "report_rows": len(report_df) if len(report_df) > 0 else 0,
                    },
                    json_mode=True,
                )
            elif not args.quiet:
                print(f"\n原始文件已更新: {result_file}")
        else:
            # Basic processing
            result_df = process_excel_files(
                str(order_file_path), str(payment_file_path), verbose=verbose
            )

            # Save result
            if args.output:
                output_path = Path(args.output)
                write_result_file(result_df, output_path)
                result_file = args.output
            else:
                write_result_file(result_df, order_file_path)
                result_file = str(order_file_path)

            # Calculate statistics for JSON output
            if args.json:
                total_rows = len(result_df)
                matched_rows = result_df["支付手续费"].notna().sum()
                match_rate = (
                    f"{(matched_rows / total_rows * 100):.2f}%"
                    if total_rows > 0
                    else "0.00%"
                )
                output_result(
                    data={
                        "output_file": result_file,
                        "statistics": {
                            "total_rows": total_rows,
                            "matched_rows": int(matched_rows),
                            "match_rate": match_rate,
                        },
                    },
                    json_mode=True,
                )
            elif not args.quiet:
                print(f"原始文件已更新: {result_file}")

        sys.exit(EXIT_SUCCESS)

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
    main()
