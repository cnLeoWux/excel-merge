import pandas as pd
import os
import re
from pathlib import Path


def extract_p_number(text):
    """
    Extract the part with "P" and following digits from a string
    """
    if pd.isna(text) or text is None:
        return None
    # Convert to string to handle numbers, then search for P pattern
    text_str = str(text)
    match = re.search(r'P\d+', text_str)
    return match.group() if match else None


def match_orders_by_p_number(external_order_no, product_name):
    """
    Match external order number with product name based on P-number
    """
    external_p = extract_p_number(external_order_no)
    product_p = extract_p_number(product_name)
    
    if external_p and product_p:
        return external_p == product_p
    return False


def read_file_with_appropriate_method(file_path):
    """
    Read a file using the appropriate pandas method based on its extension
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == '.csv':
        # For CSV files, try different encodings and parameters if default fails
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
        
        # Special handling for the second table (payment file) - ignore lines starting with #, use 5th row as header
        # Check if this is likely the payment file by looking for specific columns (商户订单号, 商品名称, etc.)
        
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
                    # Calculate header row index (0-indexed, so 4 would be 5th row)
                    header_row = skip_rows  # Use the first non-comment row as header
                    df = pd.read_csv(file_path, encoding=encoding, skiprows=skip_rows, header=0)
                else:
                    # If no comment rows, use normal parsing
                    # Try with default parameters first
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        break  # If it has expected columns, it's likely parsed correctly
                break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        
        # If the above approach failed or didn't find comment lines, use robust fallback
        if 'df' not in locals():
            for encoding in encodings:
                try:
                    # Try with default parameters first
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        break  # If it has expected columns, it's likely parsed correctly
                except (UnicodeDecodeError, pd.errors.ParserError):
                    pass
                
                try:
                    # Try with different separators
                    for sep in [',', ';', '\t']:
                        try:
                            df = pd.read_csv(file_path, encoding=encoding, sep=sep)
                            if '订单号' in df.columns or '商户订单号' in df.columns:
                                break  # If it has expected columns, it's likely parsed correctly
                        except:
                            continue
                    if 'df' in locals():
                        break
                except:
                    continue
                    
                try:
                    # Try with python engine which is more forgiving
                    df = pd.read_csv(file_path, encoding=encoding, engine='python')
                    if '订单号' in df.columns or '商户订单号' in df.columns:
                        break  # If it has expected columns, it's likely parsed correctly
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
        
        if '订单号' in df.columns:
            df['订单号'] = df['订单号'].astype(str)
        if '商户订单号' in df.columns:
            df['商户订单号'] = df['商户订单号'].astype(str)
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
        
        return pd.read_excel(file_path, dtype={'订单号': str, '商务订单号': str}, engine=engine)
    else:
        # Default to Excel reading for unknown types (as before)
        try:
            return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str}, engine='openpyxl')
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
                    return df
                except UnicodeDecodeError:
                    continue  # Try next encoding
            
            # If all encodings failed, try with utf-8-sig
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            if '订单号' in df.columns:
                df['订单号'] = df['订单号'].astype(str)
            if '商务订单号' in df.columns:
                df['商务订单号'] = df['商务订单号'].astype(str)
            return df


def process_excel_files(order_file, payment_file):
    """
    Process two files (Excel or CSV) according to the specified matching logic
    """
    # Read the files using the appropriate method
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)
    
    # Initialize the '支付手续费' column if it doesn't exist
    if '支付手续费' not in order_df.columns:
        order_df['支付手续费'] = None
    
    print("Starting matching process...")
    
    # Process each row in the order dataframe
    for idx, order_row in order_df.iterrows():
        print(f"\n--- Processing Order Row {idx} ---")
        print(f"  Full Order Number: {order_row.get('订单号', 'N/A')}")
        print(f"  External Order Number: {order_row.get('外部订单号', 'N/A')}")

        # Get order number (first 20 characters)
        original_order_no = order_row.get('订单号', '')
        if pd.isna(original_order_no) or len(str(original_order_no)) < 20:
            print(f"Row {idx}: Skipped - Order number less than 20 characters: {original_order_no}")
            continue  # Skip if order number is less than 20 characters
        order_no = str(original_order_no)[:20] 
        external_order_no = order_row.get('外部订单号', None)
        
        print(f"  Truncated Order Number (first 20 chars): {order_no}")
            
        # Determine if it's a regular order, refund order, or skip
        order_amount_raw = order_row.get('订单金额', 0)
        print(f"  Raw Order Amount: {order_amount_raw}")
        
        if pd.isna(order_amount_raw):
            order_amount = 0  # Treat NaN as 0
            print("  Order amount is NaN, setting to 0")
        else:
            # Ensure it's a numeric value to avoid issues with string values
            try:
                order_amount = float(order_amount_raw)
                print(f"  Converted Order Amount: {order_amount}")
            except (ValueError, TypeError):
                order_amount = 0  # Default to 0 if conversion fails
                print(f"  Failed to convert amount '{order_amount_raw}' to float, setting to 0")
        
        # Updated logic: positive amounts > 0 = regular order, negative amounts < 0 = refund, amount = 0 = skip
        if order_amount > 0:
            is_regular_order = True
            order_type = "正单(Regular)"
        elif order_amount < 0:
            is_regular_order = False
            order_type = "退单(Refund)"
        else:  # order_amount == 0
            print(f"Row {idx}: Skipped - Order amount is 0")
            continue  # Skip processing if amount is 0
        
        print(f"Row {idx}: Processing - Order No: {order_no}, External Order: {external_order_no}, Amount: {order_amount} ({order_type})")
        
        # Find matching records in payment dataframe
        matching_payments = []
        
        for p_idx, payment_row in payment_df.iterrows():
            print(f"    Processing Payment Row {p_idx}:")
            
            # Match by first 20 chars of order number - use only '商户订单号' as per updated requirement
            original_business_order_no = payment_row.get('商户订单号', '')
            print(f"      Payment '商户订单号': {original_business_order_no}")
            
            order_no_match = False
            p_number_match = False
            business_order_no = None  # Initialize the variable
            
            # First, try matching by first 20 characters of order number
            if pd.notna(original_business_order_no):
                original_business_str = str(original_business_order_no)
                print(f"      Length of '商户订单号': {len(original_business_str)} characters")
                
                if len(original_business_str) >= 20:
                    business_order_no = original_business_str[:20]
                    print(f"      Comparing order numbers: '{order_no}' vs '{business_order_no}'")
                    order_no_match = (order_no and business_order_no and order_no == business_order_no)
                    print(f"      Order number match result: {order_no_match}")
                else:
                    print(f"      '商户订单号' has less than 20 digits, skipping order number matching")
            else:
                print(f"      '商户订单号' is NaN, skipping order number matching")
            
            # Get product name and extract P-number (characters after the last hyphen)
            product_name = payment_row.get('商品名称', None)
            print(f"      Payment '商品名称': {product_name}")
            payment_p_number = extract_p_number(product_name)  # This gets P followed by digits
            print(f"      Extracted P-number from product name: {payment_p_number}")
            
            # For matching when order number is less than 20 chars, check after hyphen in product name
            product_after_hyphen = None
            if pd.notna(product_name):
                product_str = str(product_name)
                print(f"      Product name string: '{product_str}'")
                if '-' in product_str:
                    # Get the part after the last hyphen
                    product_after_hyphen = product_str.split('-')[-1]
                    print(f"      Part after last hyphen: '{product_after_hyphen}'")
                else:
                    print(f"      No hyphen found in product name")
            else:
                print(f"      Product name is NaN")
            
            # Check both types of matches
            print(f"      External order number: '{external_order_no}'")
            if pd.notna(external_order_no) and product_after_hyphen:
                p_number_match = str(external_order_no) == product_after_hyphen
                print(f"      Comparing external order no with product after hyphen: '{external_order_no}' == '{product_after_hyphen}' -> {p_number_match}")
            else:
                print(f"      Falling back to P-number match")
                p_number_match = match_orders_by_p_number(external_order_no, product_name)
                print(f"      P-number match result: {p_number_match}")
            
            # Get product name and extract P-number
            product_name = payment_row.get('商品名称', None)
            payment_p_number = extract_p_number(product_name)
            
           

            order_no_match = (order_no and business_order_no and order_no == business_order_no)
            p_number_match = match_orders_by_p_number(external_order_no, payment_p_number)
            
            # Check business type
            business_type = payment_row.get('业务类型', '')
            business_type_correct = (
                (is_regular_order and business_type == '收费') or 
                (not is_regular_order and business_type == '退费')
            )

             # For Observation and debugging
            print(f"  Checking payment row {p_idx}: 商户订单号: {business_order_no}, 商品名称: {product_name}, Extracted P-number: {payment_p_number}, 业务类型: {business_type}")
            
            if order_no_match:
                print(f"  - Found order number match: {order_no} == {business_order_no}")
            if p_number_match:
                # Determine which kind of match happened
                product_after_hyphen = None
                if pd.notna(product_name):
                    product_str = str(product_name)
                    if '-' in product_str:
                        product_after_hyphen = product_str.split('-')[-1]
                
                if pd.notna(external_order_no) and product_after_hyphen and str(external_order_no) == product_after_hyphen:
                    print(f"  - Found external order number after hyphen match: {external_order_no} == {product_after_hyphen}")
                else:
                    print(f"  - Found P-number match: {extract_p_number(external_order_no)} == {extract_p_number(product_name)}")
            if business_type_correct:
                print(f"  - Business type matches: {business_type} for {order_type}")
            
            match_found = ((order_no_match or p_number_match) and business_type_correct)
            print(f"    - Overall match result: ({order_no_match} OR {p_number_match}) AND {business_type_correct} = {match_found}")
            
            if match_found:
                matching_payments.append(payment_row)
                print(f"    - Match confirmed at payment row {p_idx}")
                
                # Print the matched values
                支出金额 = payment_row.get('支出金额（-元）', None)
                收入金额 = payment_row.get('收入金额（+元）', None)
                if pd.notna(支出金额):
                    print(f"    - Found 支出金额（-元）: {支出金额}")
                if pd.notna(收入金额):
                    print(f"    - Found 收入金额（+元）: {收入金额}")
        
        # If matches found, get the appropriate amount and update '支付手续费'
        if matching_payments:
            if is_regular_order:
                # For regular order, use '支出金额（-元）'
                for payment in matching_payments:
                    expenditure = payment.get('支出金额（-元）', None)
                    if expenditure is not None:
                        order_df.at[idx, '支付手续费'] = expenditure
                        print(f"  - Updated 支付手续费 for regular order: {expenditure}")
                        break  # Use first match
            else:
                # For refund order, use '收入金额（+元）'
                for payment in matching_payments:
                    income = payment.get('收入金额（+元）', None)
                    if income is not None:
                        order_df.at[idx, '支付手续费'] = income
                        print(f"  - Updated 支付手续费 for refund order: {income}")
                        break  # Use first match
        else:
            print(f"  - No matches found for this order")
    
    print("Matching process completed.")
    return order_df


def find_file_path(filename):
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
        result_df = process_excel_files(order_file_path, payment_file_path)
        
        # Modify the original order file instead of creating a new one
        original_file_extension = Path(order_file_path).suffix
        
        # Determine the appropriate engine or format based on the original file extension
        if original_file_extension.lower() == '.csv':
            result_df.to_csv(order_file_path, index=False, encoding='utf-8-sig')
        else:
            # For Excel files, determine the appropriate engine
            import zipfile
            path = Path(order_file_path)
            ext = path.suffix.lower()
            
            if ext == '.xlsx':
                try:
                    with zipfile.ZipFile(path, 'r') as zip_file:
                        engine = 'openpyxl'
                except zipfile.BadZipFile:
                    engine = 'xlrd'
                except:
                    engine = 'openpyxl'
            elif ext == '.xls':
                engine = 'xlrd'
            else:
                engine = 'openpyxl'
            
            result_df.to_excel(order_file_path, index=False, engine=engine)
        
        print(f"Original file updated: {order_file_path}")
    
    except Exception as e:
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    main()