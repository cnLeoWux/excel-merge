# Excel Merge Tool

## Overview

This is an Excel merge tool that matches two Excel/CSV files based on specific business logic. The tool processes order data from one Excel file and matches it with payment/refund details from another Excel file to update a "支付手续费" (Payment Processing Fee) column.

## Key Features

- Matches "订单号" (Order Number) with "商户订单号" (Business Order Number) using first 20 characters
- Matches "外部订单号" (External Order Number) with "商品名称" (Product Name) based on P-number patterns
- Differentiates between regular orders (positive "订单金额") and refund orders (negative "订单金额")
- Updates "支付手续费" column with appropriate values based on business type ("收费" vs "退费")
- Supports both Excel (.xlsx/.xls) and CSV file formats
- Handles various encoding issues (UTF-8, GBK, GB2312, Latin-1)
- Ignores lines starting with # in CSV files
- Modifies original files in-place instead of creating new files

## Requirements

- Python 3.7+
- Required Python packages (install with `pip install -r requirements.txt`)

## Dependencies

```bash
pandas>=1.3.0
openpyxl>=3.0.0
xlrd>=2.0.0
```

## Installation

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Interactive Mode

1. Run the script: `python excel_merge.py`
2. Enter the path/name of the first Excel file (order data) when prompted
3. Enter the path/name of the second Excel file (payment/refund data) when prompted
4. The tool will modify the original order file in-place with updated "支付手续费" values

### Command Line Interface

1. Basic usage: `python cli.py [order_file_path] [payment_file_path]`
2. Specify output file: `python cli.py [order_file_path] [payment_file_path] -o [output_file_path]`
3. The result will be saved to the specified output file or modify the original order file in-place

### Batch File (Windows)

1. Run: `run_excel_merge.bat`
2. Enter file names when prompted

## File Locations

- The tool will search for files in the current directory and the `ExcelForHandel` subdirectory
- Place your Excel files in either location to be detected by the interactive mode

## Matching Logic

1. **Order-Payment Matching**:
   - First, match "订单号" (Order Number) from the first file with "商户订单号" (Business Order Number) from the second file using the first 20 characters
   - If the order number has less than 20 characters, match "外部订单号" (External Order Number) from the first file with the part after the last "-" in "商品名称" (Product Name) from the second file
   - As an alternative, match "外部订单号" and "商品名称" based on P-number patterns (e.g., P followed by digits)

2. **Order Type Determination**:
   - Orders with positive "订单金额" (> 0) are regular orders (正单) and use "收费" business type
   - Orders with negative "订单金额" (< 0) are refund orders (退单) and use "退费" business type
   - Orders with zero "订单金额" (= 0) will have "支付手续费" set to 0.0

3. **Payment Amount Assignment**:
   - For regular orders (正单): "支付手续费" = "支出金额（-元）" from matching payment record
   - For refund orders (退单): "支付手续费" = "收入金额（+元）" from matching payment record

## Supported File Formats

- Excel files (.xlsx, .xls)
- CSV files (.csv) with various encodings (UTF-8, GBK, GB2312, Latin-1)
- CSV files with comments (lines starting with # are ignored, first non-comment line is used as header)

## Special Handling

- **CSV Comments**: Lines starting with # are ignored, and the first non-comment line is treated as the header
- **Encoding Issues**: Multiple encoding fallbacks are tried (UTF-8 → GBK → GB2312 → Latin-1 → UTF-8-sig)
- **String Preservation**: Order numbers are preserved as strings to prevent Excel's numeric conversion issues
- **NaN Handling**: Proper handling of NaN values in both order numbers and amounts
- **In-place Modification**: The original order file is modified directly instead of creating a new file

## Error Handling

- File not found errors
- Encoding and parsing errors with multiple fallback strategies
- Invalid order number length handling
- Proper exception handling with informative error messages

## Example Data Format

### Order Data (First File)

| 订单号 | 外部订单号 | 订单金额 | 支付手续费 |
|--------|------------|----------|------------|
| 40250702110303185340... | P2507021103060001 | 100.0 | (to be filled) |
| 40250701232642050749... | P2507012326430003 | -50.0 | (to be filled) |

### Payment/Refund Data (Second File)

| 商户订单号 | 商品名称 | 业务类型 | 支出金额（-元） | 收入金额（+元） |
|------------|----------|----------|------------------|------------------|
| 2025070160562408... | 吉祥旅游支付订单-P2507011154530001 | 收费 | -2.5 | (empty) |
| 2025070160570443... | 吉祥旅游支付订单-P2507011730520002 | 退费 | (empty) | 1.2 |

## Project Structure

```
excel-merge/
├── cli.py                 # Command-line interface
├── excel_merge.py         # Main implementation with interactive mode
├── README.md              # This file
├── request.md             # Original requirements document
├── requirements.txt       # Python dependencies
├── run_excel_merge.bat    # Windows batch file for easy execution
├── documents/             # Documentation files
│   ├── TECHNICAL_DOCS.md  # Technical documentation
│   ├── USAGE_EXAMPLES.md  # Usage examples
│   └── ARCHITECTURE.md    # Architecture overview
└── ExcelForHandel/        # Directory for input Excel files
```

## Documentation

- `README.md` - This file with basic usage instructions
- `documents/TECHNICAL_DOCS.md` - Technical documentation with implementation details
- `documents/USAGE_EXAMPLES.md` - Examples of how to use the tool
- `documents/ARCHITECTURE.md` - Architecture overview of the refactored codebase

## Development Notes

- Column names are in Chinese as specified in the business requirements
- The code prioritizes exact matching requirements over generalization
- Extensive logging is provided to show the matching process
- The tool tries multiple approaches to handle different file formats and encoding issues

## Troubleshooting

- If you encounter encoding errors, try saving your CSV files with UTF-8 encoding
- For Excel files with special formatting, saving as .xlsx format is recommended
- Make sure both files exist in the specified locations before running the tool
- If files are in subdirectories, provide the full path when prompted

## License

This project is open source and available under the MIT License.