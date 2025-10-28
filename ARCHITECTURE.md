# Excel Merge Tool - Architecture and Implementation Guide

## Overview

The Excel Merge Tool is designed to match two Excel/CSV files based on specific business logic. The tool processes order data from one file and matches it with payment/refund details from another file to update a "支付手续费" (Payment Processing Fee) column.

## Architecture

The refactored architecture follows a modular design to improve maintainability and eliminate code duplication:

### 1. utils.py - Shared Utilities
Contains all common functionality:
- `extract_p_number()` - Extracts P-number patterns from strings
- `match_orders_by_p_number()` - Matches orders based on P-number patterns
- `read_file_with_appropriate_method()` - Reads Excel/CSV files with encoding fallbacks
- `process_excel_files()` - Core matching algorithm
- `find_file_path()` - Locates files in various directories
- `write_result_file()` - Writes results preserving original format

### 2. excel_merge.py - Interactive Mode
- Provides interactive command-line interface
- Handles user input for file paths
- Uses shared utilities from `utils.py`

### 3. cli.py - Command-Line Interface
- Provides argument-based command-line interface
- Supports specifying output file
- Uses shared utilities from `utils.py`

## Key Improvements

### 1. Eliminated Code Duplication
- Moved all shared functionality to `utils.py`
- Both entry points now import from the common module
- File writing logic is centralized

### 2. Enhanced Maintainability
- Added type hints for better code clarity
- Improved error handling with more specific exceptions
- Better separation of concerns

### 3. Preserved Functionality
- All original matching logic remains intact
- Supports Excel (.xlsx/.xls) and CSV formats
- Handles various encodings (UTF-8, GBK, GB2312, Latin-1)
- Processes files with comment lines (starting with #)

## Matching Algorithm

The tool implements a multi-tiered matching approach:

### 1. Primary Match: Order Number vs Business Order Number
- Truncates "订单号" to first 20 characters
- Truncates "商户订单号" to first 20 characters
- Compares for equality

### 2. Secondary Match: External Order vs Product Name (Hyphen)
- When order number has less than 20 chars
- Extracts part after last "-" in "商品名称"
- Compares with "外部订单号"

### 3. Alternative Match: P-Number Pattern
- Extracts P-number (P followed by digits) from both fields
- Compares P-number patterns

### 4. Business Type Validation
- Regular orders (positive "订单金额") must match "收费" records
- Refund orders (negative "订单金额") must match "退费" records
- Zero amount orders get "支付手续费" set to 0.0

## Usage

### Interactive Mode
```bash
python excel_merge.py
```

### Command-Line Interface
```bash
python cli.py order_file.xlsx payment_file.csv
```

Optional output file:
```bash
python cli.py order_file.xlsx payment_file.csv -o result.xlsx
```

## Error Handling

- File existence validation
- Encoding fallback strategies
- Proper exception handling with meaningful messages
- NaN value handling in numeric comparisons

## Security Considerations

- Input validation for file paths
- Proper handling of file operations

## Performance Notes

The current implementation uses nested loops which results in O(n*m) complexity. For large datasets, consider implementing more efficient matching algorithms using pandas merge operations.

## Development Guidelines

For future enhancements:
1. Maintain separation of concerns
2. Use shared functions in utils.py when possible
3. Add type hints to new functions
4. Follow consistent error handling patterns
5. Ensure all functionality is tested