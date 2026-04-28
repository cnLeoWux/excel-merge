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
    
    JSON mode: outputs envelope to stdout
    Text mode: prints human-readable message to stdout
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
            print(f"Error: {error['message']}")
        
        if data:
            if "output_file" in data:
                print(f"Result saved to: {data['output_file']}")
            
            if "report_file" in data and data["report_file"]:
                print(f"Report saved to: {data['report_file']}")
            
            if "warnings" in data and data["warnings"]:
                for warning in data["warnings"]:
                    print(f"Warning: {warning}")
            
            if "message" in data:
                print(data["message"])


def main_cli():
    import argparse

    parser = argparse.ArgumentParser(
        description="Merge two Excel files based on specific matching logic."
    )
    parser.add_argument(
        "order_file", type=str, help="Path to the first Excel file (order data)"
    )
    parser.add_argument(
        "payment_file",
        type=str,
        help="Path to the second Excel file (payment/refund data)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Output filename (default: modify original file)",
    )

    # 第二阶段参数：生成销售报表
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

        # 判断是否执行第二阶段工作流
        if args.month:
            # 第二阶段：完整工作流（处理 + 标记 + 生成报表）
            if not args.json and not args.quiet:
                print(f"\n执行完整销售报表工作流...")
                print(f"  目标月份: {args.month}")

            updated_df, report_df = process_sales_report_workflow(
                args.order_file,
                args.payment_file,
                args.month,
                output_dir=args.output_dir,
                verbose=verbose,
            )

            warnings = []

            report_file = None
            if len(report_df) > 0:
                report_filename = f"report_{args.month}.xlsx"
                report_file = str(Path(args.output_dir or ".") / report_filename)
                if not args.json and not args.quiet:
                    print(f"\n新报表文件: {report_filename}")
                    print(f"包含 {len(report_df)} 行数据")
            else:
                if not args.json and not args.quiet:
                    print(f"\n没有符合条件的数据生成报表")

            output_path_str = str(args.output or args.order_file)
            try:
                if args.output:
                    output_path = Path(args.output)
                    write_result_file(updated_df, output_path)
                    if not args.json and not args.quiet:
                        print(f"\n更新后的文件已保存到: {args.output}")
                else:
                    original_file_path = Path(args.order_file)
                    write_result_file(updated_df, original_file_path)
                    if not args.json and not args.quiet:
                        print(f"\n原始文件已更新: {args.order_file}")
            except Exception as e:
                warn_msg = f"无法保存更新后的订单文件 '{output_path_str}': {str(e)}"
                warnings.append(warn_msg)
                if not args.json and not args.quiet:
                    print(f"\n警告: {warn_msg}")
                    print("月度报表已尝试生成，但原始订单文件未被更新。")

            # Calculate statistics
            total_rows = len(updated_df)
            matched_rows = updated_df["支付手续费"].notna().sum()
            match_rate = f"{(matched_rows / total_rows * 100):.2f}%" if total_rows > 0 else "0.00%"

            output_result(
                data={
                    "output_file": output_path_str,
                    "statistics": {
                        "total_rows": total_rows,
                        "matched_rows": int(matched_rows),
                        "match_rate": match_rate,
                    },
                    "report_file": report_file,
                    "report_rows": len(report_df) if len(report_df) > 0 else 0,
                    "warnings": warnings if warnings else None,
                },
                json_mode=args.json,
            )
        else:
            # 原有逻辑：只处理匹配，不执行第二阶段
            result_df = process_excel_files(
                args.order_file, args.payment_file, verbose=verbose
            )

            # If output is specified, save to that file; otherwise modify the original order file
            if args.output:
                output_path = Path(args.output)
                write_result_file(result_df, output_path)
                if not args.json:
                    output_result(
                        data={"output_file": str(args.output)},
                        json_mode=args.json,
                    )
            else:
                # Modify the original order file
                original_file_path = Path(args.order_file)
                write_result_file(result_df, original_file_path)
                if not args.json:
                    output_result(
                        data={"output_file": str(args.order_file)},
                        json_mode=args.json,
                    )

            # Calculate statistics for JSON output
            if args.json:
                total_rows = len(result_df)
                matched_rows = result_df["支付手续费"].notna().sum()
                match_rate = f"{(matched_rows / total_rows * 100):.2f}%" if total_rows > 0 else "0.00%"
                output_result(
                    data={
                        "output_file": str(args.output or args.order_file),
                        "statistics": {
                            "total_rows": total_rows,
                            "matched_rows": int(matched_rows),
                            "match_rate": match_rate,
                        },
                    },
                    json_mode=args.json,
                )

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
    main_cli()
