import pandas as pd
from pathlib import Path

# Create a sample ExcelForHandel directory if it doesn't exist
excel_dir = Path("ExcelForHandel")
excel_dir.mkdir(exist_ok=True)

# Create sample order data
order_data = {
    '订单号': ['12345678901234567890', '12345678901234567891', '12345678901234567892', '11111111111111111111'],  # 4 orders
    '外部订单号': ['ProductP123', 'ProductP456', 'ProductP789', 'ProductP000'],  # External order IDs
    '订单金额': [100.00, -50.00, 200.00, -30.00],  # Positive for regular, negative for refund
    '其他列': ['A', 'B', 'C', 'D']  # Some other data
}
order_df = pd.DataFrame(order_data)

# Create sample payment/refund data
payment_data = {
    '商户订单号': ['12345678901234567890', '12345678901234567891', '12345678901234567892', '11111111111111111111'],  # Business order IDs
    '商品名称': ['ItemP123', 'ItemP456', 'ItemP789', 'ItemP000'],  # Product names with P numbers
    '业务类型': ['收费', '退费', '收费', '退费'],  # Business type: Charge or Refund
    '支出金额（-元）': [2.50, None, 5.00, None],  # Expenditure amounts (for charges)
    '收入金额（+元）': [None, 1.25, None, 0.75],  # Income amounts (for refunds)
    '其他列': ['X', 'Y', 'Z', 'W']  # Some other data
}
payment_df = pd.DataFrame(payment_data)

# Save sample Excel files
order_df.to_excel(excel_dir / "order_sample.xlsx", index=False)
payment_df.to_excel(excel_dir / "payment_sample.xlsx", index=False)

print("Sample Excel files created in ExcelForHandel directory:")
print("- order_sample.xlsx")
print("- payment_sample.xlsx")
print("\nThese files contain test data for verifying the Excel merge functionality.")