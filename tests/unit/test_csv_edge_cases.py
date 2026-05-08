import pandas as pd
from utils import read_file_with_appropriate_method, process_excel_files
import tempfile
import os

def test_csv_edge_cases():
    with tempfile.TemporaryDirectory() as tmpdir:
        order_path = os.path.join(tmpdir, "order.csv")
        payment_path = os.path.join(tmpdir, "payment.csv")
        
        # Create an order file with long numbers and ="" artifacts
        with open(order_path, "w", encoding="utf-8-sig") as f:
            f.write("订单号,外部订单号,订单金额,订单状态\n")
            f.write('="202602123456789012345",,100.0,已确认\n')
            
        # Create a payment file with tab artifacts and a malformed row
        with open(payment_path, "w", encoding="gbk") as f:
            f.write("商户订单号,商品名称,业务类型,支出金额（-元）\n")
            f.write('\t202602123456789012345,Test,收费,-5.0\n')
            # Malformed row with too many commas
            f.write('123,456,789,0,1,2,3\n')
            
        # Process files
        df = process_excel_files(order_path, payment_path, verbose=True)
        print("RESULT:")
        print(df[["订单号", "支付手续费"]])
        
        # Assertions
        assert df is not None
        assert "支付手续费" in df.columns
        assert df.at[0, "支付手续费"] == -5.0
        assert df.at[0, "订单号"] == "202602123456789012345"

if __name__ == "__main__":
    test_csv_edge_cases()
