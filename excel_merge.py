import pandas as pd
import os
import re
from pathlib import Path
from utils import process_excel_files, read_file_with_appropriate_method, find_file_path, write_result_file


def main():
    # Get file names from user input
    print("Excel Merge Tool")
    print("Please specify the two Excel files to process:")
    
    order_input = input("Enter the path/name of the first Excel file (order data): ").strip()
    payment_input = input("Enter the path/name of the second Excel file (payment/refund data): ").strip()
    
    # Try to find the files in common locations
    order_file_path = find_file_path(order_input)
    payment_file_path = find_file_path(payment_input)
    
    # Check if files exist
    if not order_file_path.exists():
        print(f"Error: File '{order_input}' does not exist in current directory or ExcelForHandel subdirectory.")
        return
        
    if not payment_file_path.exists():
        print(f"Error: File '{payment_input}' does not exist in current directory or ExcelForHandel subdirectory.")
        return
    
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