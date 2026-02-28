# PROJECT KNOWLEDGE BASE

**Generated:** 2026-02-28
**Commit:** HEAD

## OVERVIEW
Excel Merge Tool - matches Excel/CSV files based on business logic, processing.order data with payment/refund details. Core stack: Python 3.7+, pandas 1.3+, openpyxl 3.0+, xlrd 2.0+

## STRUCTURE
```
./
├── cli.py                  # CLI interface
├── excel_merge.py          # Main implementation with interactive mode
├── utils.py               # Data processing utilities
├── documents/             # Documentation files
├── ExcelForHandel/        # Input data files for processing
├── dist/                  # Deployment artifacts
└── setup.py              # Package configuration
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Main entry (interactive) | excel_merge.py | Processes Excel/CSV files interactively |
| CLI interface | cli.py | Command line wrapper for processing |
| Business logic | utils.py | Core order-payment matching algorithms |
| Documentation | documents/* | Tech docs, usage examples, architecture overview |
| Test files | *_test.py, test_*.py | Various test files scattered in root |
| Input samples | ExcelForHandel/ | Place Excel/CSV files for processing |

## CONVENTIONS
- Chinese column names for business requirements (订单号, 商户订单号, etc.)
- File detection in current directory and ExcelForHandel/ subdirectory
- In-place modification of original order file
- Supports multiple encodings (UTF-8, GBK, GB2312, Latin-1) for CSV files

## ANTI-PATTERNS (THIS PROJECT)
- Files mixed in root directory instead of src/ structure
- Tests placed in root instead of tests/ directory
- No __init__.py files for proper Python package structure

## UNIQUE STYLES
- CSV comment handling (lines starting with # ignored)
- String preservation to prevent Excel numeric conversion issues
- Multiple matching strategies based on order number length 

## COMMANDS
```bash
# Interactive mode
python excel_merge.py

# CLI mode
python cli.py [order_file_path] [payment_file_path]

# With output redirection
python cli.py [order_file_path] [payment_file_path] -o [output_file_path]

# Install dependencies
pip install -r requirements.txt
```

## NOTES
- ExcelForHandel directory used for default file location in interactive mode
- Matches "订单号" with "商户订单号" using first 20 characters
- Differentiates between regular orders (positive amount = "收费") and refunds (negative = "退费")