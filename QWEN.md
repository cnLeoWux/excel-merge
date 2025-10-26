# QWEN.md

## Project Overview

This is an **Excel Merge Tool** project that implements functionality to merge two Excel files based on specific business logic. The tool processes order data from one Excel file and matches it with payment/refund details from another Excel file to update a "支付手续费" (Payment Processing Fee) column.

### Key Features
- Matches "订单号" (Order Number) with "商务订单号" (Business Order Number) using first 20 characters
- Matches "外部订单号" (External Order Number) with "商品名称" (Product Name) based on P-number patterns
- Differentiates between regular orders (positive "订单金额") and refund orders (negative "订单金额")
- Updates "支付手续费" column with appropriate values based on business type ("收费" vs "退费")

## Project Structure

```
D:\Workspace\excel-merge\
├── cli.py                 # Command-line interface for the tool
├── create_sample_data.py  # Script to create sample Excel files for testing
├── excel_merge.py         # Main implementation of the Excel merge logic
├── README.md             # Project documentation
├── request.md            # Original requirements document
├── requirements.txt      # Python dependencies
├── run_excel_merge.bat   # Windows batch file for easy execution
├── verify_original.py    # Script to verify original data
├── verify_result.py      # Script to verify merged results
└── ExcelForHandel/       # Directory for input Excel files
```

## Building and Running

### Prerequisites
- Python 3.7+
- Required Python packages (install with `pip install -r requirements.txt`)

### Installation
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running the Tool
- **Interactive mode:**
  ```bash
  python excel_merge.py
  ```
  This will prompt you to enter the file paths/names for both Excel files. The tool will search for the files in the current directory and the ExcelForHandel subdirectory.
  
- **Command-line interface:**
  ```bash
  python cli.py [order_file_path] [payment_file_path]
  ```
  Optionally specify output file:
  ```bash
  python cli.py [order_file_path] [payment_file_path] -o [output_file_path]
  ```
  
- **Windows batch file:**
  ```bash
  run_excel_merge.bat
  ```
  Note: The batch file runs the interactive version, so you'll need to provide file names when prompted.

The tool will generate a result file named `merged_result_[original_file_name].xlsx` in the current directory (unless specified otherwise).

## Development Conventions

- The code is written in Python using pandas for Excel file processing
- Column names in Chinese are used as specified in the requirements
- String data types are preserved for order numbers to prevent Excel numeric conversion issues
- The tool enforces exactly two Excel files in the processing directory
- Error handling is implemented for common issues like missing directories or files

## Key Files

- `excel_merge.py`: Contains the core matching and merging logic
- `cli.py`: Provides command-line interface with configurable directory
- `request.md`: Original specification document with exact requirements
- `README.md`: User-facing documentation with usage instructions
- `requirements.txt`: Lists required Python packages (pandas, openpyxl)