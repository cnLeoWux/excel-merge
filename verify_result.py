import pandas as pd

# Read the merged result to verify the logic
result_df = pd.read_excel("ExcelForHandel/merged_result_order_sample.xlsx")
print("Merged Result Data:")
print(result_df)
print("\nColumn names in the merged result:")
print(list(result_df.columns))

# Verify the logic is correctly applied
print("\nVerification:")
for idx, row in result_df.iterrows():
    order_amount = row['订单金额']
    payment_fee = row['支付手续费'] if '支付手续费' in result_df.columns else 'N/A'
    print(f"Row {idx}: Order Amount = {order_amount}, Payment Fee = {payment_fee}")