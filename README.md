# excel-merge

This tool is designed to merge two Excel files based on a specific matching logic:

1. The first Excel file contains order data
2. The second Excel file contains payment and refund details

## How It Works

- Matches the "订单号" (Order Number) from the first file with the "商务订单号" (Business Order Number) from the second file (first 20 characters)
- Matches the "外部订单号" (External Order Number) from the first file with the "商品名称" (Product Name) from the second file based on the P-number pattern
- For regular orders (positive "订单金额"), extracts "支出金额（-元）" from records where "业务类型" is "收费"
- For refund orders (negative "订单金额"), extracts "收入金额（+元）" from records where "业务类型" is "退费"
- Updates the "支付手续费" column in the first file with the matched values

## Usage

### Interactive Mode:
1. Run the script: `python excel_merge.py`
2. Enter the path/name of the first Excel file (order data) when prompted
3. Enter the path/name of the second Excel file (payment/refund data) when prompted
4. The result will be saved as `merged_result_[original_file_name].xlsx` in the current directory

### Command Line Interface:
1. Run: `python cli.py [order_file_path] [payment_file_path]`
2. Optionally specify output file: `python cli.py [order_file_path] [payment_file_path] -o [output_file_path]`
3. The result will be saved to the specified output file or as `merged_result_[original_file_name].xlsx`

## Dependencies

Install the required packages using:
```
pip install -r requirements.txt
```