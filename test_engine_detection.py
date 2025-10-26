import pandas as pd
from pathlib import Path
import zipfile


def get_excel_engine(file_path):
    """
    Determine the appropriate Excel engine for a given file
    """
    # First, check the file extension
    path = Path(file_path)
    ext = path.suffix.lower()
    
    # If it's .xlsx, try to verify it's actually a zip file
    if ext == '.xlsx':
        try:
            with zipfile.ZipFile(path, 'r') as zip_file:
                # If it's a valid zip file, use openpyxl
                return 'openpyxl'
        except zipfile.BadZipFile:
            # If it's not a valid zip file, fall back to xlrd (sometimes older xls files have xlsx extension)
            return 'xlrd'
        except:
            # For any other error, fall back to openpyxl
            return 'openpyxl'
    elif ext == '.xls':
        # For .xls files, use xlrd
        return 'xlrd'
    else:
        # Default to openpyxl for unknown extensions
        return 'openpyxl'


# Test with sample file paths
order_file_path = "ExcelForHandel/order_sample.xlsx"
payment_file_path = "ExcelForHandel/payment_sample.xlsx"

print(f"Order file: {order_file_path}, determined engine: {get_excel_engine(order_file_path)}")
print(f"Payment file: {payment_file_path}, determined engine: {get_excel_engine(payment_file_path)}")

# Try to read the files with the determined engines
order_engine = get_excel_engine(order_file_path)
payment_engine = get_excel_engine(payment_file_path)

order_df = pd.read_excel(order_file_path, dtype={'订单号': str}, engine=order_engine)
payment_df = pd.read_excel(payment_file_path, dtype={'商户订单号': str}, engine=payment_engine)

print("Successfully read both Excel files with the appropriate engines!")
print("Order DataFrame shape:", order_df.shape)
print("Payment DataFrame shape:", payment_df.shape)
print("Order DataFrame columns:", list(order_df.columns))
print("Payment DataFrame columns:", list(payment_df.columns))