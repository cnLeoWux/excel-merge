# Project Summary

## Overall Goal
Create an Excel merge tool that matches two Excel/CSV files based on specific business logic: matching "订单号" with "商务订单号" (first 20 characters), "外部订单号" with "商品名称" (P-number pattern), and updating "支付手续费" based on order type (正单/退单) and business type (收费/退费).

## Key Knowledge
- **Technology Stack**: Python with pandas, openpyxl, xlrd for Excel/CSV processing
- **File Locations**: Process files from `ExcelForHandel` directory
- **Column Names**: 
  - Order file: '订单号', '外部订单号', '订单金额', '支付手续费'
  - Payment file: '商务订单号', '商品名称', '业务类型', '支出金额（-元）', '收入金额（+元）'
- **Matching Logic**:
  - Orders with positive '订单金额' are regular orders (正单), negative are refunds (退单)
  - Regular orders match payment records with '收费' business type, get '支出金额（-元）'
  - Refund orders match payment records with '退费' business type, get '收入金额（+元）'
- **Special Requirements**:
  - Ignore lines starting with # in payment CSV files
  - Use first non-comment line as header (typically 5th line)
  - Modify original file in-place instead of creating new files
  - Provide detailed matching process output

## Recent Actions
1. **[DONE]** Implemented core matching logic with order number (20char) and P-number matching
2. **[DONE]** Added support for both Excel (.xlsx/.xls) and CSV file formats
3. **[DONE]** Created interactive and CLI interfaces for specifying file names
4. **[DONE]** Added robust error handling for encoding issues (UTF-8, GBK, GB2312, Latin-1)
5. **[DONE]** Implemented CSV comment line handling (lines starting with #)
6. **[DONE]** Added detailed matching process output to console
7. **[DONE]** Fixed "object of type 'float' has no len()" error by adding NaN checks
8. **[DONE]** Implemented in-place file modification instead of creating new files
9. **[DONE]** Added support for various CSV parsing issues with multiple fallback strategies
10. **[DONE]** Created comprehensive test files to verify functionality

## Current Plan
1. **[DONE]** Complete the Excel merge tool with all specified requirements
2. **[DONE]** Ensure robust handling of various file formats and edge cases
3. **[DONE]** Implement error handling for encoding and parsing issues
4. **[DONE]** Final testing with sample data to confirm all functionality works correctly
5. **[DONE]** Document the solution in requirements format

---

## Summary Metadata
**Update time**: 2025-10-26T14:14:05.119Z 
