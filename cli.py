import pandas as pd
import os
import re
from pathlib import Path
import argparse
from utils import process_excel_files, read_file_with_appropriate_method, find_file_path, write_result_file


def main_cli():
    parser = argparse.ArgumentParser(description='Merge two Excel files based on specific matching logic.')
    parser.add_argument('order_file', type=str, help='Path to the first Excel file (order data)')
    parser.add_argument('payment_file', type=str, help='Path to the second Excel file (payment/refund data)')
    parser.add_argument('-o', '--output', type=str, default=None, help='Output filename (default: modify original file)')
    
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
        result_df = process_excel_files(args.order_file, args.payment_file, verbose=True)
        
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


if __name__ == "__main__":
    main_cli()