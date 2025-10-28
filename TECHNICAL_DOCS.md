# Excel Merge Tool - Technical Documentation

## Overview

The Excel Merge Tool is a Python application that performs complex matching between two Excel/CSV files based on specific business logic. The application implements matching between order data and payment/refund details to update a payment processing fee column.

## Architecture

The system consists of two primary files:

1. `excel_merge.py` - Contains the main implementation with interactive mode
2. `cli.py` - Provides command-line interface functionality

Both files share the same core processing logic but differ in their user interface approach.

## Core Components

### 1. File Reading Functions

The `read_file_with_appropriate_method()` function handles reading Excel and CSV files with sophisticated error handling:

- **Excel Support**: Handles both .xlsx and .xls formats with appropriate engines (openpyxl for .xlsx, xlrd for .xls)
- **CSV Support**: Tries multiple encodings (UTF-8, GBK, GB2312, Latin-1) with fallback strategies
- **Comment Handling**: Special processing for CSV files with comment lines starting with #
- **Type Preservation**: Ensures order numbers are read as strings to prevent numeric conversion

### 2. Matching Logic Functions

The matching logic uses two functions:

- `extract_p_number(text)`: Extracts P followed by digits from a string (e.g., P2507021103060001)
- `match_orders_by_p_number(external_order_no, product_name)`: Matches external order numbers with product names based on P-number patterns

### 3. Main Processing Function

The `process_excel_files()` function implements the core business logic:

- Reads both files using the appropriate method
- Iterates through each order in the order file
- Determines if each order is a regular order (positive amount) or refund order (negative amount)
- Matches with payment records using multiple matching strategies
- Updates the "支付手续费" column with appropriate values

## Matching Algorithm

### Primary Match: Order Number vs Business Order Number

1. Truncate the "订单号" to first 20 characters
2. Truncate the "商户订单号" to first 20 characters
3. Compare for equality

### Secondary Match: External Order Number vs Product Name

When the order number has less than 20 characters:

1. Extract the part after the last "-" in "商品名称"
2. Compare with "外部订单号"

### Alternative Match: P-Number Pattern

As an additional matching strategy:

1. Extract P-number (P followed by digits) from "外部订单号"
2. Extract P-number from "商品名称"
3. Compare for equality

### Business Type Validation

- Regular orders (positive "订单金额") must match "收费" records
- Refund orders (negative "订单金额") must match "退费" records
- Zero amount orders get "支付手续费" set to 0.0

## Data Processing Logic

### Regular Orders (正单)
- When "订单金额" > 0
- Use "支出金额（-元）" from matching payment records
- Apply to "支付手续费" column

### Refund Orders (退单)
- When "订单金额" < 0
- Use "收入金额（+元）" from matching payment records
- Apply to "支付手续费" column

### Zero Amount Orders
- When "订单金额" == 0
- Set "支付手续费" to 0.0 without further matching

## Error Handling & Fallbacks

### File Reading Fallbacks
1. Try different encodings sequentially
2. Use different CSV separators if standard comma fails
3. Use Python engine for more forgiving parsing
4. Skip bad lines if parsing continues to fail

### Data Validation
1. Check for NaN values in critical fields
2. Validate string lengths before substring operations
3. Convert numeric values safely with try-catch blocks

### Matching Validation
1. Log each step of the matching process
2. Continue processing if no match is found rather than crashing
3. Use first valid match when multiple matches exist

## Input/Output Handling

### File Location Support
- Supports files in current directory
- Supports files in ExcelForHandel subdirectory
- Accepts full file paths

### In-place File Modification
- Rather than creating new files, modifies the original order file
- Preserves original file format (Excel or CSV)
- Uses appropriate engines based on file extension

## Command-Line Interface

The `cli.py` file provides a command-line interface with:

- Argument parsing using argparse
- Optional output file specification
- Input file validation before processing
- Flexible file path support

## Performance Considerations

1. **Memory Usage**: Loads entire files into pandas DataFrames for processing
2. **Time Complexity**: O(n*m) where n is number of order records and m is number of payment records
3. **Optimizations**: Could implement indexing for large files, but current approach is suitable for typical business use cases

## Testing & Verification

The tool provides detailed console output during matching showing:
- Each order being processed
- Comparison values being used
- Match results
- Final payment fee assignments

## Known Limitations

1. For very large files, memory usage can be substantial
2. Multiple matches use only the first valid match
3. CSV file format assumes standard structure after comment lines

## Future Enhancements

1. Add support for additional file formats
2. Implement performance optimizations for large files
3. Add more detailed logging with configurable levels
4. Improve user experience with progress bars for large datasets