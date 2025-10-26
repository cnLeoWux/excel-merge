import pandas as pd
from pathlib import Path
import traceback
import re

def read_file_with_appropriate_method(file_path):
    """
    Read a file using the appropriate pandas method based on its extension
    """
    from pathlib import Path
    import zipfile
    
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == '.csv':
        # For CSV files, dtype for specific columns is not supported in older pandas versions
        # So we'll read it normally and convert the specific columns to string afterwards
        df = pd.read_csv(file_path)
        if '订单号' in df.columns:
            df['订单号'] = df['订单号'].astype(str)
        if '商户订单号' in df.columns:
            df['商户订单号'] = df['商户订单号'].astype(str)
        return df
    elif ext in ['.xlsx', '.xls']:
        # For Excel files
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
        
        return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str}, engine=engine)
    else:
        # Default to Excel reading for unknown types (as before)
        try:
            return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str}, engine='openpyxl')
        except:
            # Try reading as CSV if Excel reading fails
            df = pd.read_csv(file_path)
            if '订单号' in df.columns:
                df['订单号'] = df['订单号'].astype(str)
            if '商户订单号' in df.columns:
                df['商户订单号'] = df['商户订单号'].astype(str)
            return df


def extract_p_number(text):
    """
    Extract the part with "P" and following digits from a string
    """
    if pd.isna(text) or text is None:
        return None
    match = re.search(r'P\d+', str(text))
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


def process_excel_files_debug(order_file, payment_file):
    """
    Process two files (Excel or CSV) according to the specified matching logic
    """
    # Read the files using the appropriate method
    print("Reading order file...")
    order_df = read_file_with_appropriate_method(order_file)
    print("Reading payment file...")
    payment_df = read_file_with_appropriate_method(payment_file)
    
    print("Order DataFrame shape:", order_df.shape)
    print("Payment DataFrame shape:", payment_df.shape)
    
    # Initialize the '支付手续费' column if it doesn't exist
    if '支付手续费' not in order_df.columns:
        order_df['支付手续费'] = None
    
    print("Starting to process rows...")
    # Process each row in the order dataframe
    for idx, order_row in order_df.iterrows():
        # Get order number (first 20 characters)
        order_no = str(order_row.get('订单号', ''))[:20] if pd.notna(order_row.get('订单号')) else None
        external_order_no = order_row.get('外部订单号', None)
        
        # Skip if order number is less than 20 characters
        if order_no and len(str(order_row.get('订单号', ''))) < 20:
            continue
            
        # Determine if it's a regular order or refund order
        order_amount = order_row.get('订单金额', 0)
        is_regular_order = order_amount >= 0
        
        print(f"Processing row {idx}: Order Amount = {order_amount}, Regular Order = {is_regular_order}")
        
        # Find matching records in payment dataframe
        matching_payments = []
        
        for p_idx, payment_row in payment_df.iterrows():
            # Match by first 20 chars of order number
            business_order_no = str(payment_row.get('商户订单号', ''))[:20] if pd.notna(payment_row.get('商户订单号')) else None
            
            # Match by P-number in external order no and product name
            product_name = payment_row.get('商品名称', None)
            
            order_no_match = (order_no and business_order_no and order_no == business_order_no)
            p_number_match = match_orders_by_p_number(external_order_no, product_name)
            
            # Check business type
            business_type = payment_row.get('业务类型', '')
            business_type_correct = (
                (is_regular_order and business_type == '收费') or 
                (not is_regular_order and business_type == '退费')
            )
            
            if ((order_no_match or p_number_match) and business_type_correct):
                matching_payments.append(payment_row)
        
        # If matches found, get the appropriate amount and update '支付手续费'
        if matching_payments:
            print(f"Found {len(matching_payments)} matching payments for row {idx}")
            if is_regular_order:
                # For regular order, use '支出金额（-元）'
                for payment in matching_payments:
                    expenditure = payment.get('支出金额（-元）', None)
                    if expenditure is not None:
                        order_df.at[idx, '支付手续费'] = expenditure
                        print(f"Set 支付手续费 for regular order: {expenditure}")
                        break  # Use first match
            else:
                # For refund order, use '收入金额（+元）'
                for payment in matching_payments:
                    income = payment.get('收入金额（+元）', None)
                    if income is not None:
                        order_df.at[idx, '支付手续费'] = income
                        print(f"Set 支付手续费 for refund order: {income}")
                        break  # Use first match
    
    print("Processing complete, returning result")
    return order_df


try:
    result_df = process_excel_files_debug("ExcelForHandel/order.csv", "ExcelForHandel/payment.csv")
    print("Saving result...")
    output_filename = "debug_test_result.xlsx"
    result_df.to_excel(output_filename, index=False)
    print(f"Result saved to: {output_filename}")
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()