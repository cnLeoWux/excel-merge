# Excel Merge Tool - Usage Examples

## Overview

This document provides detailed examples of how to use the Excel Merge Tool, including sample data formats, command-line usage, and expected output.

## Sample Data Formats

### Order File (First File) - Sample.xlsx
```
订单号                    外部订单号           订单金额    支付手续费
4025070211030318534001   P2507021103060001   100.00
4025070123264205074902   P2507012326430003   -50.00
4025070922460638851403   P2507092246080005   0.00
```

### Payment/Refund File (Second File) - Payments.csv
```
# This is a comment line
# Another comment
# Third comment
# Fourth comment
商户订单号                商品名称                             业务类型  支出金额（-元）  收入金额（+元）
202507016056240801       吉祥旅游支付订单-P2507011154530001    收费      -2.50
202507016057044302       吉祥旅游支付订单-P2507011730520002    退费                  1.20
202507016057868103       吉祥旅游支付订单-P2507092246080005    收费      -1.00
```

## Expected Matching Process

### Example 1: Regular Order Match
- Order: 订单号=4025070211030318534001, 外部订单号=P2507021103060001, 订单金额=100.00 (正单)
- Payment: 商户订单号=202507016056240801, 商品名称=吉祥旅游支付订单-P2507011154530001, 业务类型=收费
- Process:
  1. Extract first 20 chars: 40250702110303185340 == 202507016056240801 (No match)
  2. Extract P-number: P2507021103060001 vs P2507011154530001 (No match)
  3. No match found - 支付手续费 remains empty or set to 0

### Example 2: Refund Order Match
- Order: 订单号=4025070123264205074902, 外部订单号=P2507012326430003, 订单金额=-50.00 (退单)
- Payment: 商户订单号=202507016057044302, 商品名称=吉祥旅游支付订单-P2507011730520002, 业务类型=退费
- Process:
  1. Extract first 20 chars: 40250701232642050749 == 202507016057044302 (No match)
  2. Extract P-number: P2507012326430003 vs P2507011730520002 (No match)
  3. No match found - 支付手续费 remains empty or set to 0

## Command-Line Usage Examples

### Basic Usage
```bash
# Process two files and modify the first file in-place
python cli.py order_data.xlsx payment_data.xlsx

# Process two files and save to a specific output file
python cli.py order_data.xlsx payment_data.xlsx -o result.xlsx

# Process CSV files
python cli.py order_data.csv payment_data.csv

# Process with mixed file types
python cli.py order_data.xlsx payment_data.csv
```

### Interactive Mode Usage
```bash
# Run in interactive mode
python excel_merge.py

# Follow the prompts:
# Enter the path/name of the first Excel file (order data): order_data.xlsx
# Enter the path/name of the second Excel file (payment/refund data): payment_data.xlsx
```

### File Location Examples
```bash
# Files in current directory
python cli.py order_data.xlsx payment_data.xlsx

# Files in ExcelForHandel subdirectory (interactive mode will find these)
python excel_merge.py
# Enter: order_data.xlsx
# Enter: payment_data.xlsx

# Full paths
python cli.py /full/path/to/order_data.xlsx /full/path/to/payment_data.xlsx

# Relative paths
python cli.py ./data/orders.xlsx ./data/payments.xlsx
```

## Processing Output Example

When running the tool, you'll see output like:

```
Starting matching process...

--- Processing Order Row 0 ---
  Full Order Number: 4025070211030318534001
  External Order Number: P2507021103060001
  Truncated Order Number (first 20 chars): 40250702110303185340
  Raw Order Amount: 100.0
  Converted Order Amount: 100.0
Row 0: Processing - Order No: 40250702110303185340, External Order: P2507021103060001, Amount: 100.0 (正单(Regular))
    Processing Payment Row 0:
      Payment '商户订单号': 202507016056240801
      Length of '商户订单号': 18 characters
      '商户订单号' has less than 20 digits, skipping order number matching
      Payment '商品名称': 吉祥旅游支付订单-P2507011154530001
      Extracted P-number from product name: P2507011154530001
      Part after last hyphen: P2507011154530001
      External order number: 'P2507021103060001'
      Comparing external order no with product after hyphen: 'P2507021103060001' == 'P2507011154530001' -> False
      Falling back to P-number match
      P-number match result: False
  Checking payment row 0: 商户订单号: None, 商品名称: 吉祥旅游支付订单-P2507011154530001, Extracted P-number: P2507011154530001, 业务类型: 收费
    - Business type matches: 收费 for 正单(Regular)
    - Overall match result: (False OR False) AND True = False
    - No matches found for this order
...
Matching process completed.
Original file updated: order_data.xlsx
```

## Encoding and Format Examples

### UTF-8 CSV
```
订单号,外部订单号,订单金额
4025070211030318534001,P2507021103060001,100.00
```

### GBK CSV (Chinese characters)
```
订单号,外部订单号,订单金额
4025070211030318534001,P2507021103060001,100.00
```

### CSV with Comments
```
# This file contains order data
# Generated on 2023-01-01
# Do not modify manually
# Header starts on next line
订单号,外部订单号,订单金额
4025070211030318534001,P2507021103060001,100.00
```

## Common Scenarios and Expected Results

### Scenario 1: Successful Match
- Input: Order with 订单号=4025070211030318534001, 订单金额=100.00 (正单)
- Payment: Record with 商户订单号=4025070211030318534002 (first 20 chars match), 业务类型=收费
- Result: 支付手续费 = 支出金额（-元） from matching payment record

### Scenario 2: Refund Order Match
- Input: Order with 订单号=4025070123264205074902, 订单金额=-50.00 (退单)
- Payment: Record with 商户订单号=4025070123264205074903 (first 20 chars match), 业务类型=退费
- Result: 支付手续费 = 收入金额（+元） from matching payment record

### Scenario 3: Zero Amount Order
- Input: Order with 订单金额=0.00
- Result: 支付手续费 automatically set to 0.00, no matching performed

### Scenario 4: No Match Found
- Input: Order with no corresponding payment record
- Result: 支付手续费 remains unchanged (or set to null if not present)

## Troubleshooting Examples

### Error: File Not Found
```
Error: File 'order_data.xlsx' does not exist in current directory or ExcelForHandel subdirectory.
```
**Solution**: Verify the file exists in the current directory or ExcelForHandel subdirectory, or provide the full path.

### Error: Encoding Issues
```
Error processing files: 'utf-8' codec can't decode byte 0xa1 in position 0: invalid start byte
```
**Solution**: The tool automatically tries different encodings (GBK, GB2312, Latin-1) as fallbacks.

### Error: Invalid Data Types
```
Error processing files: could not convert string to float: 'N/A'
```
**Solution**: Ensure numeric columns (like 订单金额) contain valid numeric values or empty cells.