import pandas as pd
from pathlib import Path

# Test the engine detection function
def get_engine(file_path):
    if file_path.suffix.lower() == '.xlsx':
        return 'openpyxl'
    elif file_path.suffix.lower() == '.xls':
        return 'xlrd'
    else:
        # Default to openpyxl for .xlsx, xlrd for .xls
        if '.' not in file_path.name:
            return 'openpyxl'  # default to xlsx format
        return 'openpyxl'  # default to openpyxl for .xlsx files

# Test with sample file paths
order_file_path = Path("ExcelForHandel/order_sample.xlsx")
payment_file_path = Path("ExcelForHandel/payment_sample.xlsx")

print(f"Order file: {order_file_path}, engine: {get_engine(order_file_path)}")
print(f"Payment file: {payment_file_path}, engine: {get_engine(payment_file_path)}")

# Try to read the files with the determined engines
order_engine = get_engine(order_file_path)
payment_engine = get_engine(payment_file_path)

order_df = pd.read_excel(order_file_path, dtype={'订单号': str}, engine=order_engine)
payment_df = pd.read_excel(payment_file_path, dtype={'商户订单号': str}, engine=payment_engine)

print("Successfully read both Excel files with the appropriate engines!")
print("Order DataFrame shape:", order_df.shape)
print("Payment DataFrame shape:", payment_df.shape)