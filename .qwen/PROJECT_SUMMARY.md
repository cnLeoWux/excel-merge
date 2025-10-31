# Project Summary

## Overall Goal
Create and refactor an Excel merge tool that matches two Excel/CSV files based on specific business logic to update a "支付手续费" column, with robust file format support, proper error handling, and well-organized documentation.

## Key Knowledge
- **Technology Stack**: Python with pandas, openpyxl, xlrd for Excel/CSV processing
- **Directory Structure**: 
  - Main files: `cli.py`, `excel_merge.py`, `utils.py` (shared utilities)
  - Documentation in: `documents/` directory
  - Requirements: `requirements.txt` (pandas>=1.3.0, openpyxl>=3.0.0, xlrd>=2.0.0)
- **Column Names**: 
  - Order file: '订单号', '外部订单号', '订单金额', '支付手续费'
  - Payment file: '商户订单号', '商品名称', '业务类型', '支出金额（-元）', '收入金额（+元）'
- **Matching Logic**:
  - Orders with positive '订单金额' are regular orders (正单), negative are refunds (退单), zero amounts get '支付手续费' set to 0.0
  - Regular orders match payment records with '收费' business type, get '支出金额（-元）'
  - Refund orders match payment records with '退费' business type, get '收入金额（+元）'
  - Primary match: First 20 characters of '订单号' with '商户订单号'
  - Secondary match: '外部订单号' with part after last "-" in '商品名称'
  - Alternative match: P-number pattern matching
- **Architecture**: Modular design with shared utilities in `utils.py` to eliminate code duplication between `cli.py` and `excel_merge.py`
- **File Formats**: Supports both Excel (.xlsx/.xls) and CSV files with multiple encoding fallbacks (UTF-8, GBK, GB2312, Latin-1)
- **Special Requirements**: Handle CSV files with comment lines (starting with #), preserve original file format when updating, modify files in-place rather than creating new files

## Recent Actions
- [DONE] Created shared `utils.py` module to eliminate code duplication between `excel_merge.py` and `cli.py`
- [DONE] Refactored core functionality into shared utility functions: `extract_p_number`, `match_orders_by_p_number`, `read_file_with_appropriate_method`, `process_excel_files`, `find_file_path`, `write_result_file`
- [DONE] Added comprehensive type hints throughout the codebase for better maintainability
- [DONE] Improved error handling with specific exception types and proper file handling
- [DONE] Organized documentation into a dedicated `documents/` directory with `TECHNICAL_DOCS.md`, `USAGE_EXAMPLES.md`, and `ARCHITECTURE.md`
- [DONE] Updated `README.md` to reflect new documentation location and project structure
- [DONE] Maintained all original matching logic while improving performance and maintainability
- [DONE] Tested refactored code with sample data to confirm functionality
- [DONE] Renamed original requirements document to `REQUIREMENT.md`
- [DONE] Committed and pushed all changes to the remote repository

## Current Plan
1. [DONE] Complete refactoring of codebase to eliminate duplication
2. [DONE] Add type hints and improve code quality
3. [DONE] Create comprehensive documentation in organized directory structure
4. [DONE] Update README to reflect new architecture and documentation
5. [DONE] Test refactored code with sample data
6. [DONE] Commit and push all changes to remote repository
7. [DONE] Verify repository is properly updated with all changes

---

## Summary Metadata
**Update time**: 2025-10-31T06:25:17.345Z 
