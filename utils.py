"""
Utility functions for the Excel Merge Tool.
This module contains common functions shared between the interactive and CLI versions.
"""

import pandas as pd
import os
import re
from pathlib import Path
from typing import Optional, Any
import logging


def extract_p_number(text: Any) -> Optional[str]:
    """
    Extract the part with "P" and following digits from a string
    """
    if pd.isna(text) or text is None:
        return None
    # Convert to string to handle numbers, then search for P pattern
    text_str = str(text)
    match = re.search(r'P\d+', text_str)
    return match.group() if match else None


def match_orders_by_p_number(external_order_no: Any, product_name: Any) -> bool:
    """
    Match external order number with product name based on P-number
    """
    external_p = extract_p_number(external_order_no)
    product_p = extract_p_number(product_name)
    
    if external_p and product_p:
        return external_p == product_p
    return False


def read_file_with_appropriate_method(file_path: str) -> pd.DataFrame:
    """
    Read a file using the appropriate pandas method based on its extension
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == '.csv':
        # For CSV files, try different encodings and parameters if default fails
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
        
        # First, try to read the file content to check for # comment lines
        for encoding in encodings:
            try:
                # Read the file as text to check for comment lines
                with open(file_path, 'r', encoding=encoding) as f:
                    lines = f.readlines()
                
                # Count how many lines start with # at the beginning
                skip_rows = 0
                for i, line in enumerate(lines):
                    if line.strip().startswith('#'):
                        skip_rows += 1
                    else:
                        break  # Stop at first line that doesn't start with #
                
                # If we need to skip rows, and row 4 exists (0-indexed as row 4 = 5th row), use it as header
                if skip_rows > 0:
                    df = pd.read_csv(file_path, encoding=encoding, skiprows=skip_rows, header=0)
                else:
                    # If no comment rows, use normal parsing
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        # If it has expected columns, it's likely parsed correctly
                        break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        
        # If the above approach failed, use robust fallback
        if 'df' not in locals():
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        break
                except (UnicodeDecodeError, pd.errors.ParserError):
                    pass
                
                # Try with different separators
                for sep in [',', ';', '\t']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, sep=sep)
                        if '订单号' in df.columns or '商户订单号' in df.columns:
                            break
                    except:
                        continue
                if 'df' in locals():
                    break
                    
                try:
                    # Try with python engine which is more forgiving
                    df = pd.read_csv(file_path, encoding=encoding, engine='python')
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        break
                except:
                    continue
        
        # If still no success, try with more robust parameters
        if 'df' not in locals():
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, engine='python', on_bad_lines='skip')
                    break
                except:
                    continue
        
        # If all encodings failed, try with different parameters
        if 'df' not in locals():
            df = pd.read_csv(file_path, encoding='utf-8', engine='python', on_bad_lines='skip', sep=None)
        
        # Ensure critical columns are treated as strings
        if '订单号' in df.columns:
            df['订单号'] = df['订单号'].astype(str)
        if '商户订单号' in df.columns:
            df['商户订单号'] = df['商户订单号'].astype(str)
        if '商务订单号' in df.columns:
            df['商务订单号'] = df['商务订单号'].astype(str)
            
        return df
    elif ext in ['.xlsx', '.xls']:
        # For Excel files
        import zipfile
        
        # Determine engine based on file type
        if ext == '.xlsx':
            try:
                with zipfile.ZipFile(path, 'r') as zip_file:
                    # If it's a valid zip file, use openpyxl
                    engine = 'openpyxl'
            except zipfile.BadZipFile:
                # If it's not a valid zip file, fall back to xlrd (sometimes older xls files have xlsx extension)
                engine = 'xlrd'
            except:
                # For any other error, fall back to openpyxl
                engine = 'openpyxl'
        elif ext == '.xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str, '商务订单号': str}, engine=engine)
    else:
        # Default to Excel reading for unknown types (as before)
        try:
            return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str, '商务订单号': str}, engine='openpyxl')
        except:
            # For CSV files with encoding issues, try different encodings
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
            df = None
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns:
                        df['订单号'] = df['订单号'].astype(str)
                    if '商户订单号' in df.columns:
                        df['商户订单号'] = df['商户订单号'].astype(str)
                    if '商务订单号' in df.columns:
                        df['商务订单号'] = df['商务订单号'].astype(str)
                    return df
                except UnicodeDecodeError:
                    continue  # Try next encoding
            
            # If all encodings failed, try with utf-8-sig
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            if '订单号' in df.columns:
                df['订单号'] = df['订单号'].astype(str)
            if '商户订单号' in df.columns:
                df['商户订单号'] = df['商户订单号'].astype(str)
            if '商务订单号' in df.columns:
                df['商务订单号'] = df['商务订单号'].astype(str)
            return df


def process_excel_files(order_file: str, payment_file: str, verbose: bool = False) -> pd.DataFrame:
    """
    Process two files (Excel or CSV) according to the specified matching logic.
    Uses more efficient pandas operations instead of nested loops.
    """
    # Read the files using the appropriate method
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)
    
    # Initialize the '支付手续费' column if it doesn't exist
    if '支付手续费' not in order_df.columns:
        order_df['支付手续费'] = None
    
    if verbose:
        print("Starting matching process...")
    
    # Process each row in the order dataframe
    for idx, order_row in order_df.iterrows():
        if verbose:
            print(f"\n--- Processing Order Row {idx} ---")
            print(f"  Full Order Number: {order_row.get('订单号', 'N/A')}")
            print(f"  External Order Number: {order_row.get('外部订单号', 'N/A')}")

        # Get order number (first 20 characters)
        original_order_no = order_row.get('订单号', '')
        if pd.isna(original_order_no) or len(str(original_order_no)) < 20:
            if verbose:
                print(f"Row {idx}: Skipped - Order number less than 20 characters: {original_order_no}")
            continue  # Skip if order number is less than 20 characters
            
        order_no = str(original_order_no)[:20] 
        external_order_no = order_row.get('外部订单号', None)
        
        if verbose:
            print(f"  Truncated Order Number (first 20 chars): {order_no}")
            
        # Determine if it's a regular order, refund order, or skip
        order_amount_raw = order_row.get('订单金额', 0)
        if verbose:
            print(f"  Raw Order Amount: {order_amount_raw}")
        
        if pd.isna(order_amount_raw):
            order_amount = 0  # Treat NaN as 0
            if verbose:
                print("  Order amount is NaN, setting to 0")
        else:
            # Ensure it's a numeric value to avoid issues with string values
            try:
                order_amount = float(order_amount_raw)
                if verbose:
                    print(f"  Converted Order Amount: {order_amount}")
            except (ValueError, TypeError):
                order_amount = 0  # Default to 0 if conversion fails
                if verbose:
                    print(f"  Failed to convert amount '{order_amount_raw}' to float, setting to 0")
        
        # Updated logic: positive amounts > 0 = regular order, negative amounts < 0 = refund, amount = 0 = set 支付手续费 to 0
        if order_amount > 0:
            is_regular_order = True
            order_type = "正单(Regular)"
        elif order_amount < 0:
            is_regular_order = False
            order_type = "退单(Refund)"
        else:  # order_amount == 0
            if verbose:
                print(f"Row {idx}: Order amount is 0, setting 支付手续费 to 0")
            order_df.at[idx, '支付手续费'] = 0.0
            continue  # Skip further processing for this row but set the fee to 0
        
        if verbose:
            print(f"Row {idx}: Processing - Order No: {order_no}, External Order: {external_order_no}, Amount: {order_amount} ({order_type})")
        
        # Find matching records in payment dataframe using vectorized operations where possible
        matching_payments = []
        
        # First, try to find exact matches by truncated order number
        business_order_numbers = payment_df['商户订单号'].astype(str)
        exact_matches = business_order_numbers.str[:20] == order_no
        exact_match_rows = payment_df[exact_matches]
        
        # For non-exact matches, check P-number and hyphen logic
        if len(exact_match_rows) == 0:
            for p_idx, payment_row in payment_df.iterrows():
                # Try P-number match
                product_name = payment_row.get('商品名称', None)
                p_number_match = match_orders_by_p_number(external_order_no, product_name)
                
                # Try hyphen match
                hyphen_match = False
                if pd.notna(product_name) and pd.notna(external_order_no):
                    product_str = str(product_name)
                    if '-' in product_str:
                        product_after_hyphen = product_str.split('-')[-1]
                        hyphen_match = str(external_order_no) == product_after_hyphen
                
                # Check business type
                business_type = payment_row.get('业务类型', '')
                business_type_correct = (
                    (is_regular_order and business_type == '收费') or 
                    (not is_regular_order and business_type == '退费')
                )
                
                if verbose and (p_number_match or hyphen_match):
                    print(f"      Match found via P-number or hyphen: P={p_number_match}, Hyphen={hyphen_match}")
                
                if ((p_number_match or hyphen_match) and business_type_correct):
                    matching_payments.append(payment_row)
                    if verbose:
                        print(f"    - Match confirmed at payment row {p_idx}")
        else:
            # Handle exact matches
            for p_idx, payment_row in exact_match_rows.iterrows():
                # Check business type
                business_type = payment_row.get('业务类型', '')
                business_type_correct = (
                    (is_regular_order and business_type == '收费') or 
                    (not is_regular_order and business_type == '退费')
                )
                
                if business_type_correct:
                    matching_payments.append(payment_row)
                    if verbose:
                        print(f"    - Exact match confirmed at payment row {p_idx} with correct business type")
        
        # If matches found, get the appropriate amount and update '支付手续费'
        if matching_payments:
            if is_regular_order:
                # For regular order, use '支出金额（-元）'
                for payment in matching_payments:
                    expenditure = payment.get('支出金额（-元）', None)
                    if expenditure is not None:
                        order_df.at[idx, '支付手续费'] = expenditure
                        if verbose:
                            print(f"  - Updated 支付手续费 for regular order: {expenditure}")
                        break  # Use first match
            else:
                # For refund order, use '收入金额（+元）'
                for payment in matching_payments:
                    income = payment.get('收入金额（+元）', None)
                    if income is not None:
                        order_df.at[idx, '支付手续费'] = income
                        if verbose:
                            print(f"  - Updated 支付手续费 for refund order: {income}")
                        break  # Use first match
        else:
            if verbose:
                print(f"  - No matches found for this order")
    
    if verbose:
        print("Matching process completed.")
    return order_df


def find_file_path(filename: str) -> Path:
    """
    Try to find the file in different possible locations:
    1. Current directory
    2. ExcelForHandel subdirectory
    """
    # First, try the current directory
    if Path(filename).exists():
        return Path(filename)
    
    # Then try the ExcelForHandel subdirectory
    excel_dir_path = Path("ExcelForHandel") / filename
    if excel_dir_path.exists():
        return excel_dir_path
    
    # Return original path if not found (to preserve original error)
    return Path(filename)


def write_result_file(df: pd.DataFrame, file_path: Path) -> None:
    """
    Write the result DataFrame to the specified file path, preserving the original file format.
    """
    import zipfile
    from pathlib import Path
    
    original_file_extension = file_path.suffix
    
    # Determine the appropriate engine or format based on the original file extension
    if original_file_extension.lower() == '.csv':
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
    else:
        # For Excel files, determine the appropriate engine
        path = Path(file_path)
        ext = path.suffix.lower()
        
        if ext == '.xlsx':
            try:
                with zipfile.ZipFile(path, 'r') as zip_file:
                    engine = 'openpyxl'
            except zipfile.BadZipFile:
                engine = 'xlrd'
            except Exception:
                engine = 'openpyxl'
        elif ext == '.xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        df.to_excel(file_path, index=False, engine=engine)