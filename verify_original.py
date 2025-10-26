import pandas as pd

# Read the original files to verify the matching logic
order_df = pd.read_excel("ExcelForHandel/order_sample.xlsx")
payment_df = pd.read_excel("ExcelForHandel/payment_sample.xlsx")

print("Original Order Data:")
print(order_df)
print("\nOriginal Payment/Refund Data:")
print(payment_df)

print("\nLogic Verification:")
for idx, order_row in order_df.iterrows():
    order_amount = order_row['订单金额']
    order_type = "正单(Regular)" if order_amount >= 0 else "退单(Refund)"
    print(f"Row {idx}: Amount = {order_amount} ({order_type})")