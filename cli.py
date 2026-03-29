import pandas as pd
import os
import re
from pathlib import Path
import argparse
from utils import (
    process_excel_files,
    read_file_with_appropriate_method,
    find_file_path,
    write_result_file,
    add_sales_report_period,
    filter_unmarked_and_generate_report,
    process_sales_report_workflow,
)


def main_cli():
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

    args = parser.parse_args()

    # Check if files exist
    if not Path(args.order_file).exists():
        print(f"Error: File '{args.order_file}' does not exist.")
        return

    if not Path(args.payment_file).exists():
        print(f"Error: File '{args.payment_file}' does not exist.")
        return

    print(f"Processing files:")
    print(f"  Order file: {args.order_file}")
    print(f"  Payment/Refund file: {args.payment_file}")

    try:
        # 判断是否执行第二阶段工作流
        if args.month:
            # 第二阶段：完整工作流（处理 + 标记 + 生成报表）
            print(f"\n执行完整销售报表工作流...")
            print(f"  目标月份: {args.month}")

            updated_df, report_df = process_sales_report_workflow(
                args.order_file,
                args.payment_file,
                args.month,
                output_dir=args.output_dir,
                verbose=True,
            )

            # 保存更新后的原文件
            if args.output:
                output_path = Path(args.output)
                write_result_file(updated_df, output_path)
                print(f"\n更新后的文件已保存到: {args.output}")
            else:
                original_file_path = Path(args.order_file)
                write_result_file(updated_df, original_file_path)
                print(f"\n原始文件已更新: {args.order_file}")

            # 报告生成的报表
            if len(report_df) > 0:
                report_filename = f"report_{args.month}.xlsx"
                print(f"\n新报表文件: {report_filename}")
                print(f"包含 {len(report_df)} 行数据")
            else:
                print(f"\n没有符合条件的数据生成报表")
        else:
            # 原有逻辑：只处理匹配，不执行第二阶段
            result_df = process_excel_files(
                args.order_file, args.payment_file, verbose=True
            )

            # If output is specified, save to that file; otherwise modify the original order file
            if args.output:
                output_path = Path(args.output)
                write_result_file(result_df, output_path)
                print(f"Result saved to: {args.output}")
            else:
                # Modify the original order file
                original_file_path = Path(args.order_file)
                write_result_file(result_df, original_file_path)
                print(f"Original file updated: {args.order_file}")

    except Exception as e:
        print(f"Error processing files: {e}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main_cli()
