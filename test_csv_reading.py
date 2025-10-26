import pandas as pd
from pathlib import Path


def read_file_with_appropriate_method(file_path):
    """
    Read a file using the appropriate pandas method based on its extension
    """
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


# Test the function
print("Testing CSV file reading:")
try:
    order_df = read_file_with_appropriate_method("ExcelForHandel/order.csv")
    payment_df = read_file_with_appropriate_method("ExcelForHandel/payment.csv")
    print("Order DataFrame shape:", order_df.shape)
    print("Payment DataFrame shape:", payment_df.shape)
    print("Successfully read CSV files!")
    print("Order DataFrame:")
    print(order_df)
    print("\nPayment DataFrame:")
    print(payment_df)
except Exception as e:
    print(f"Error reading files: {e}")