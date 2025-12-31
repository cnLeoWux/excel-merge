'''Author: Leo Wu leo.wux@lego.com
Date: 2025-10-27 10:59:13
LastEditors: Leo Wu leo.wux@lego.com
LastEditTime: 2025-12-30 13:23:11
FilePath: /excel-merge/excel_merge.py
Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'''
import pandas as pd
import os
import re
from pathlib import Path
from utils import process_excel_files, read_file_with_appropriate_method, find_file_path, write_result_file


def main():
    # Get file names from user input
    print("Excel Merge Tool")
    
    # Get all files in ExcelForHandel directory
    excel_dir = Path("ExcelForHandel")
    if not excel_dir.exists():
        print(f"Error: ExcelForHandel directory does not exist.")
        return
    
    # List all files in the directory
    files = list(excel_dir.glob("*"))
    if not files:
        print(f"Error: No files found in ExcelForHandel directory.")
        return
    
    # Display files for selection
    print("Available files in ExcelForHandel directory:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file.name}")
    
    # Get first file selection
    while True:
        try:
            order_choice = int(input("\nSelect the first Excel file (order data) by number: ").strip())
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
            payment_choice = int(input("\nSelect the second Excel file (payment/refund data) by number: ").strip())
            if 1 <= payment_choice <= len(files):
                payment_file_path = files[payment_choice - 1]
                break
            else:
                print(f"Please enter a number between 1 and {len(files)}.")
        except ValueError:
            print("Please enter a valid number.")
    
    print(f"Processing files:")
    print(f"  Order file: {order_file_path}")
    print(f"  Payment/Refund file: {payment_file_path}")
    
    try:
        result_df = process_excel_files(str(order_file_path), str(payment_file_path), verbose=True)
        
        # Modify the original order file instead of creating a new one
        write_result_file(result_df, order_file_path)
        
        print(f"Original file updated: {order_file_path}")
    
    except Exception as e:
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()
