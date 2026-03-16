# PROJECT KNOWLEDGE BASE

**Generated:** 2026-03-16

## OVERVIEW
Excel Merge Tool - matches Excel/CSV files based on business logic, processing order data with payment/refund details. Core stack: Python 3.7+, pandas, openpyxl, xlrd.

## STRUCTURE
```
./
├── cli.py                  # CLI interface (argparse)
├── excel_merge.py          # Interactive mode entry point
├── excel_merge_api.py      # Flask API wrapper
├── utils.py               # Core business logic (~500 lines)
├── setup.py              # Package config with console_scripts
├── requirements.txt      # Dependencies
├── openspec/            # OpenSpec configuration
├── documents/          # Documentation files
├── ExcelForHandel/    # Input data directory
└── test_*.py          # Test files (4 files in root)
```

## BUILD / LINT / TEST COMMANDS

### Run Tests
```bash
# Run all tests
python -m pytest

# Run single test file
python -m pytest test_engine.py
python -m pytest test_csv_reading.py
python -m pytest test_problematic_csv.py
python -m pytest test_engine_detection.py

# Run with verbose output
python -m pytest -v

# Run specific test function (if tests use pytest structure)
python -m pytest test_engine.py::test_function_name -v

# Alternative: run test files directly
python test_engine.py
python test_csv_reading.py
```

### Install Dependencies
```bash
pip install -r requirements.txt

# Install as editable package
pip install -e .
```

### Run Application
```bash
# Interactive mode
python excel_merge.py

# CLI mode
python cli.py order.xlsx payment.xlsx
python cli.py order.xlsx payment.xlsx -o result.xlsx

# Flask API mode (if implemented)
python excel_merge_api.py

# Console scripts (after pip install)
excel-merge          # Interactive mode
excel-merge-cli      # CLI mode
```

## CODE STYLE GUIDELINES

### Imports
- **Order**: Standard library → Third-party → Local modules
- **Example**:
  ```python
  import os
  import re
  from pathlib import Path
  from typing import Optional, Any

  import pandas as pd

  from utils import process_excel_files, find_file_path
  ```
- Always use absolute imports for local modules

### Naming Conventions
- **Functions**: `snake_case` (e.g., `process_excel_files`, `extract_p_number`)
- **Variables**: `snake_case` (e.g., `order_df`, `file_path`)
- **Constants**: `UPPER_CASE` (rarely used)
- **Module names**: `snake_case` (e.g., `utils.py`, `cli.py`)
- **Chinese column names** are used for business requirements (订单号, 商户订单号, etc.)

### Formatting
- **Indentation**: 4 spaces (no tabs)
- **Line length**: ~100-120 characters (no strict limit observed)
- **Blank lines**: 2 lines between module-level functions
- **Docstrings**: Triple quotes for module and function documentation

### Type Hints
- Use type hints for function parameters and return types
- Common types: `Optional[str]`, `Optional[Any]`, `pd.DataFrame`, `Path`
- Example: `def extract_p_number(text: Any) -> Optional[str]:`

### Error Handling
- Use specific exceptions: `UnicodeDecodeError`, `pd.errors.ParserError`, `ValueError`, `TypeError`
- Provide informative error messages with context
- Use try/except blocks for file I/O and data parsing
- Avoid bare `except:` clauses (use `except Exception:` if needed)
- Print errors to stdout for user feedback in CLI mode

### Functions
- Keep functions focused and single-purpose
- Use docstrings explaining purpose and parameters
- Verbose flag pattern for optional debug output: `verbose: bool = False`
- Return `Optional[T]` or `None` for functions that may fail gracefully

### DataFrame Handling
- Force string type for order number columns: `astype(str)`
- Use pandas NA checking: `pd.isna()`, `pd.notna()`
- Preserve original file format on write (CSV vs Excel detection)

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Interactive mode | excel_merge.py | Lists files in ExcelForHandel/, prompts for selection |
| CLI interface | cli.py | argparse wrapper, supports -o for output path |
| Business logic | utils.py | Order-payment matching, P-number extraction, encoding handling |
| API server | excel_merge_api.py | Flask wrapper for HTTP API |
| Entry points | setup.py | console_scripts: excel-merge, excel-merge-cli |
| Tests | test_*.py | 4 test files in root (not in tests/ dir) |
| OpenSpec config | openspec/ | project.md, config.yaml |

## CODE MAP

### Main Symbols (utils.py)
| Symbol | Purpose |
|--------|---------|
| `process_excel_files()` | Main orchestration, iterates orders, calls matchers |
| `extract_p_number()` | Regex extract P+digits pattern |
| `match_orders_by_p_number()` | Match external_order_no with product_name |
| `read_file_with_appropriate_method()` | Handles CSV/Excel with encoding fallback |
| `write_result_file()` | Preserves original format on write |
| `find_file_path()` | Searches current dir then ExcelForHandel/ |

### Entry Points
| Script | Function | Behavior |
|--------|----------|----------|
| excel_merge.py | `main()` | Interactive file selection from ExcelForHandel/ |
| cli.py | `main_cli()` | CLI args: order_file, payment_file, optional -o output |
| excel_merge_api.py | Flask app | HTTP API for file processing |

## CONVENTIONS

### Business Logic
- **Order matching**: First 20 chars of "订单号" ↔ "商户订单号"
- **P-number matching**: Extract P+digits from external_order_no and product_name
- **Hyphen matching**: Match external_order_no with part after last "-" in product_name
- **Order type by amount**: positive=正单(charge), negative=退单(refund), zero=skip
- **Encoding priority**: gbk → utf-8 → gb2312 → latin-1 → utf-8-sig
- **Type preservation**: Order numbers forced to string to prevent Excel numeric conversion
- **CSV comments**: Lines starting with # ignored, first non-comment line = header
- **In-place modification**: Original order file modified directly (not creating new file)
- **File discovery**: Searches current directory → ExcelForHandel/ subdirectory

### Dependencies
```
pandas>=1.3.0
openpyxl>=3.0.0
xlrd>=2.0.0
flask>=2.0.0
werkzeug>=2.0.0
```

## ANTI-PATTERNS (THIS PROJECT)
- Files mixed in root directory instead of src/ structure
- Tests placed in root instead of tests/ directory
- No __init__.py files (not a proper Python package)
- No .github/workflows or CI configuration
- No linting/formatting config (flake8, black, etc.)
- No pytest.ini or pyproject.toml for test configuration
- No type checking with mypy

## UNIQUE STYLES
- **Multiple engine detection**: zipfile check for xlsx vs xls engine selection
- **Column name flexibility**: Search for columns containing substrings (e.g., '商户' + '订单')
- **Verbose logging**: Optional verbose flag prints detailed matching progress
- **Encoding fallback chain**: Try multiple encodings before failing
