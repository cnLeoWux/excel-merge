"""
Utility functions for the Excel Merge Tool.
This module contains common functions shared between the interactive and CLI versions.
"""

import pandas as pd
import os
import re
from pathlib import Path
from typing import Optional, Any
import logging


def extract_p_number(text: Any) -> Optional[str]:
    """
    Extract the part with "P" and following digits from a string
    """
    if pd.isna(text) or text is None:
        return None
    # Convert to string to handle numbers, then search for P pattern
    text_str = str(text)
    match = re.search(r'P\d+', text_str)
    return match.group() if match else None


def match_orders_by_p_number(external_order_no: Any, product_name: Any) -> bool:
    """
    Match external order number with product name based on P-number
    """
    external_p = extract_p_number(external_order_no)
    product_p = extract_p_number(product_name)
    
    if external_p and product_p:
        return external_p == product_p
    return False


def read_file_with_appropriate_method(file_path: str) -> pd.DataFrame:
    """
    Read a file using the appropriate pandas method based on its extension
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == '.csv':
        # For CSV files, try different encodings and parameters if default fails
        # Prioritize gbk encoding for Chinese files
        encodings = ['gbk', 'utf-8', 'gb2312', 'latin-1', 'utf-8-sig']
        
        df = None
        
        # First, try to read the file with proper comment handling
        for encoding in encodings:
            try:
                # Read the file as text to check for # comment lines
                with open(file_path, 'r', encoding=encoding) as f:
                    lines = f.readlines()
                
                # Count how many lines start with # at the beginning
                skip_rows = 0
                for line in lines:
                    if line.strip().startswith('#'):
                        skip_rows += 1
                    else:
                        break  # Stop at first line that doesn't start with #
                
                # Read the CSV file with proper skiprows and encoding
                df = pd.read_csv(
                    file_path, 
                    encoding=encoding, 
                    skiprows=skip_rows, 
                    header=0,  # Use the first non-comment line as header
                    engine='python',  # Use python engine for better encoding handling
                    on_bad_lines='skip'  # Skip bad lines
                )
                
                # Check if we have at least some expected columns or reasonable data
                if df.shape[0] > 0 and df.shape[1] > 5:
                    break  # We have a valid dataframe
                    
            except UnicodeDecodeError:
                continue
            except pd.errors.ParserError:
                continue
            except Exception as e:
                print(f"Error reading CSV file with encoding {encoding}: {e}")
                continue
        
        # If still no success, try with different separators
        if df is None or df.shape[0] == 0:
            for encoding in encodings:
                for sep in [',', ';', '\t']:
                    try:
                        df = pd.read_csv(
                            file_path, 
                            encoding=encoding, 
                            sep=sep, 
                            engine='python',
                            on_bad_lines='skip'
                        )
                        if df.shape[0] > 0 and df.shape[1] > 5:
                            break
                    except:
                        continue
                if df is not None and df.shape[0] > 0:
                    break
        
        # If all else fails, try with automatic detection
        if df is None or df.shape[0] == 0:
            df = pd.read_csv(
                file_path, 
                encoding='gbk',  # Default to gbk for Chinese files
                engine='python',
                on_bad_lines='skip',
                sep=None,  # Auto-detect separator
                skip_blank_lines=True
            )
        
        # Ensure critical columns are treated as strings
        for col in df.columns:
            col_str = str(col)
            if '订单' in col_str or '流水' in col_str:
                df[col] = df[col].astype(str)
                
        return df
    elif ext in ['.xlsx', '.xls']:
        # For Excel files
        import zipfile
        
        # Determine engine based on file type
        if ext == '.xlsx':
            try:
                with zipfile.ZipFile(path, 'r') as zip_file:
                    # If it's a valid zip file, use openpyxl
                    engine = 'openpyxl'
            except zipfile.BadZipFile:
                # If it's not a valid zip file, fall back to xlrd (sometimes older xls files have xlsx extension)
                engine = 'xlrd'
            except:
                # For any other error, fall back to openpyxl
                engine = 'openpyxl'
        elif ext == '.xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str, '商务订单号': str}, engine=engine)
    else:
        # Default to Excel reading for unknown types (as before)
        try:
            return pd.read_excel(file_path, dtype={'订单号': str, '商户订单号': str, '商务订单号': str}, engine='openpyxl')
        except:
            # For CSV files with encoding issues, try different encodings
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
            df = None
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if '订单号' in df.columns:
                        df['订单号'] = df['订单号'].astype(str)
                    if '商户订单号' in df.columns:
                        df['商户订单号'] = df['商户订单号'].astype(str)
                    if '商务订单号' in df.columns:
                        df['商务订单号'] = df['商务订单号'].astype(str)
                    return df
                except UnicodeDecodeError:
                    continue  # Try next encoding
            
            # If all encodings failed, try with utf-8-sig
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            if '订单号' in df.columns:
                df['订单号'] = df['订单号'].astype(str)
            if '商户订单号' in df.columns:
                df['商户订单号'] = df['商户订单号'].astype(str)
            if '商务订单号' in df.columns:
                df['商务订单号'] = df['商务订单号'].astype(str)
            return df


def process_excel_files(order_file: str, payment_file: str, verbose: bool = False) -> pd.DataFrame:
    """
    Process two files (Excel or CSV) according to the specified matching logic.
    Uses more efficient pandas operations instead of nested loops.
    """
    # Read the files using the appropriate method
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)
    
    # Initialize the '支付手续费' column if it doesn't exist
    if '支付手续费' not in order_df.columns:
        order_df['支付手续费'] = None
    
    if verbose:
        print("Starting matching process...")
        print(f"Order file columns: {list(order_df.columns)}")
        print(f"Payment file columns: {list(payment_df.columns)}")
    
    # Process each row in the order dataframe
    for idx, order_row in order_df.iterrows():
        if verbose:
            print(f"\n--- Processing Order Row {idx} ---")
            print(f"  Full Order Number: {order_row.get('订单号', 'N/A')}")
            print(f"  External Order Number: {order_row.get('外部订单号', 'N/A')}")
            print(f"  Initial Payment Fee: {order_row.get('支付手续费', 'N/A')}")

        # Get order number (first 20 characters)
        original_order_no = order_row.get('订单号', '')
        order_no = str(original_order_no)[:20] if pd.notna(original_order_no) else ''
        
        # Don't skip orders with short order numbers, continue for P-number matching
        if pd.isna(original_order_no) or len(str(original_order_no)) < 20:
            if verbose:
                print(f"Row {idx}: Order number less than 20 characters: {original_order_no}, continuing for P-number matching")
        external_order_no = order_row.get('外部订单号', None)
        
        if verbose:
            print(f"  Truncated Order Number (first 20 chars): {order_no}")
            
        # Determine if it's a regular order, refund order, or skip
        order_amount_raw = order_row.get('订单金额', 0)
        if verbose:
            print(f"  Raw Order Amount: {order_amount_raw}")
        
        if pd.isna(order_amount_raw):
            order_amount = 0  # Treat NaN as 0
            if verbose:
                print("  Order amount is NaN, setting to 0")
        else:
            # Ensure it's a numeric value to avoid issues with string values
            try:
                order_amount = float(order_amount_raw)
                if verbose:
                    print(f"  Converted Order Amount: {order_amount}")
            except (ValueError, TypeError):
                order_amount = 0  # Default to 0 if conversion fails
                if verbose:
                    print(f"  Failed to convert amount '{order_amount_raw}' to float, setting to 0")
        
        # Updated logic: positive amounts > 0 = regular order, negative amounts < 0 = refund, amount = 0 = set 支付手续费 to 0
        if order_amount > 0:
            is_regular_order = True
            order_type = "正单(Regular)"
        elif order_amount < 0:
            is_regular_order = False
            order_type = "退单(Refund)"
        else:  # order_amount == 0
            if verbose:
                print(f"Row {idx}: Order amount is 0, setting 支付手续费 to 0")
            order_df.at[idx, '支付手续费'] = 0.0
            continue  # Skip further processing for this row but set the fee to 0
        
        if verbose:
            print(f"Row {idx}: Processing - Order No: {order_no}, External Order: {external_order_no}, Amount: {order_amount} ({order_type})")
        
        # Find matching records in payment dataframe using vectorized operations where possible
        matching_payments = []
        
        # First, try to find exact matches by truncated order number
        # Check if '商户订单号' column exists, if not, try other possible column names
        business_order_col = None
        for col in payment_df.columns:
            col_str = str(col)
            if '商户' in col_str and '订单' in col_str:
                business_order_col = col
                break
        
        if not business_order_col:
            # If no column with '商户订单号' found, try other possible column names
            for col in payment_df.columns:
                if '订单' in str(col):
                    business_order_col = col
                    break
        
        if business_order_col:
            business_order_numbers = payment_df[business_order_col].astype(str)
            exact_matches = business_order_numbers.str[:20] == order_no
            exact_match_rows = payment_df[exact_matches]
            
            if verbose and len(exact_match_rows) > 0:
                print(f"  Found {len(exact_match_rows)} exact matches for order {order_no}")
        else:
            # If no column with '订单' found, skip exact match check
            exact_match_rows = payment_df.head(0)  # Empty dataframe
            if verbose:
                print(f"  No business order column found in payment file")
        
        # For non-exact matches, check P-number and hyphen logic
        if len(exact_match_rows) == 0:
            for p_idx, payment_row in payment_df.iterrows():
                # Try P-number match
                product_name = payment_row.get('商品名称', '')
                p_number_match = False
                
                # Enhanced P-number matching
                p_number_match = match_orders_by_p_number(external_order_no, product_name)
                if verbose and p_number_match:
                    external_p = extract_p_number(external_order_no)
                    product_p = extract_p_number(product_name)
                    print(f"      P-number match: {external_p} == {product_p}")
                
                # Enhanced hyphen match
                hyphen_match = False
                if pd.notna(product_name) and pd.notna(external_order_no):
                    external_str = str(external_order_no)
                    product_str = str(product_name)
                    
                    if '-' in product_str:
                        # Try matching with different separators
                        external_parts = external_str.split()
                        product_parts = product_str.split('-')
                        
                        for external_part in external_parts:
                            if external_part.strip() == product_parts[-1].strip():
                                hyphen_match = True
                                if verbose:
                                    print(f"      Hyphen match: {external_part.strip()} == {product_parts[-1].strip()}")
                                break
                
                # Check business type
                business_type = payment_row.get('业务类型', '')
                business_type_str = str(business_type).strip()
                
                # Check if business type matches expected values
                is_charge = False
                is_refund = False
                
                if '收费' in business_type_str:
                    is_charge = True
                elif '退费' in business_type_str:
                    is_refund = True
                elif '退款' in business_type_str:
                    is_refund = True
                elif '服务费' in business_type_str:
                    is_charge = True
                
                # Determine if business type is correct
                business_type_correct = False
                if is_regular_order:
                    # For regular orders, we need a '收费' type
                    if is_charge:
                        business_type_correct = True
                else:
                    # For refund orders, we need a '退费' or '退款' type
                    if is_refund:
                        business_type_correct = True
                
                if verbose:
                    print(f"      Checking payment row {p_idx}: type='{business_type_str}', charge={is_charge}, refund={is_refund}, correct={business_type_correct}")
                
                # If we have a match based on P-number or hyphen AND correct business type, add to matching payments
                if (p_number_match or hyphen_match) and business_type_correct:
                    matching_payments.append(payment_row)
                    if verbose:
                        print(f"    - Match confirmed at payment row {p_idx}")
                elif (p_number_match or hyphen_match) and not business_type_correct:
                    if verbose:
                        print(f"    - Skipped match at payment row {p_idx} (incorrect business type: {business_type_str})")
        else:
            # Handle exact matches
            for p_idx, payment_row in exact_match_rows.iterrows():
                # Check business type before adding to matching payments
                business_type = payment_row.get('业务类型', '')
                business_type_str = str(business_type).strip()
                
                # Check if this payment has the right business type for the order
                is_payment_charge = '收费' in business_type_str or '服务费' in business_type_str
                is_payment_refund = '退费' in business_type_str or '退款' in business_type_str
                
                # For regular orders, we only want charge type payments that have non-zero expenditure
                if is_regular_order and is_payment_charge:
                    expenditure = payment_row.get('支出金额（-元）', 0)
                    try:
                        expenditure_val = float(expenditure)
                        if expenditure_val != 0:
                            # Regular order (positive amount) matches with charge type and non-zero expenditure
                            matching_payments.append(payment_row)
                            if verbose:
                                print(f"    - Added exact match at payment row {p_idx} (charge type with non-zero expenditure)")
                    except (ValueError, TypeError):
                        if verbose:
                            print(f"    - Skipped exact match at payment row {p_idx} (invalid expenditure value)")
                # For refund orders, we only want refund type payments that have non-zero income
                elif not is_regular_order and is_payment_refund:
                    income = payment_row.get('收入金额（+元）', 0)
                    try:
                        income_val = float(income)
                        if income_val != 0:
                            # Refund order (negative amount) matches with refund type and non-zero income
                            matching_payments.append(payment_row)
                            if verbose:
                                print(f"    - Added exact match at payment row {p_idx} (refund type with non-zero income)")
                    except (ValueError, TypeError):
                        if verbose:
                            print(f"    - Skipped exact match at payment row {p_idx} (invalid income value)")
                else:
                    # Skip payments with incorrect business type
                    if verbose:
                        print(f"    - Skipped exact match at payment row {p_idx} (incorrect business type: {business_type_str})")
        
        # If matches found, get the appropriate amount and update '支付手续费'
        if matching_payments:
            # Find the appropriate amount columns
            income_col = '收入金额（+元）'
            expenditure_col = '支出金额（-元）'
            
            if verbose:
                print(f"  Using income column: {income_col}, expenditure column: {expenditure_col}")
            
            updated = False
            
            # For all matching payments, find the one with correct business type and amount
            for payment in matching_payments:
                business_type = payment.get('业务类型', '')
                business_type_str = str(business_type).strip()
                
                # Check if this payment has the right business type for the order
                is_payment_charge = '收费' in business_type_str or '服务费' in business_type_str
                is_payment_refund = '退费' in business_type_str or '退款' in business_type_str
                
                if is_regular_order:
                    # For regular order: 支付手续费 = 支出金额（-元）的值且需要转成负值，代表需要付出去的钱
                    if is_payment_charge:
                        expenditure = payment.get(expenditure_col, 0)
                        try:
                            expenditure_val = float(expenditure)
                            if expenditure_val != 0:
                                # 支出金额已经是负数，我们需要保持其负值状态
                                order_df.at[idx, '支付手续费'] = expenditure_val
                                if verbose:
                                    print(f"  - Updated 支付手续费 for regular order: {expenditure_val}")
                                updated = True
                                break
                        except (ValueError, TypeError):
                            if verbose:
                                print(f"  - Invalid expenditure value: {expenditure}")
                else:
                    # For refund order: 支付手续费 = 收入金额（+元）的值且应该为正数，代表退单后需要收回来的钱
                    if is_payment_refund:
                        income = payment.get(income_col, 0)
                        try:
                            income_val = float(income)
                            if income_val != 0:
                                # 收入金额已经是正数，我们需要保持其正数状态
                                order_df.at[idx, '支付手续费'] = income_val
                                if verbose:
                                    print(f"  - Updated 支付手续费 for refund order: {income_val}")
                                updated = True
                                break
                        except (ValueError, TypeError):
                            if verbose:
                                print(f"  - Invalid income value: {income}")
            
            # If no valid payment found, do not update the payment fee
            if not updated and verbose:
                print(f"  - No valid payment amount found for this order")
        else:
            if verbose:
                print(f"  - No matches found for this order")
    
    if verbose:
        print("Matching process completed.")
    
    # 添加"销售报表账期"列
    order_df = add_sales_report_period(order_df, verbose=verbose)
    
    return order_df


def find_file_path(filename: str) -> Path:
    """
    Try to find the file in different possible locations:
    1. Current directory
    2. ExcelForHandel subdirectory
    """
    # First, try the current directory
    if Path(filename).exists():
        return Path(filename)
    
    # Then try the ExcelForHandel subdirectory
    excel_dir_path = Path("ExcelForHandel") / filename
    if excel_dir_path.exists():
        return excel_dir_path
    
    # Return original path if not found (to preserve original error)
    return Path(filename)


def write_result_file(df: pd.DataFrame, file_path: Path) -> None:
    """
    Write the result DataFrame to the specified file path, preserving the original file format.
    """
    import zipfile
    from pathlib import Path
    
    original_file_extension = file_path.suffix
    
    # Determine the appropriate engine or format based on the original file extension
    if original_file_extension.lower() == '.csv':
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
    else:
        # For Excel files, determine the appropriate engine
        path = Path(file_path)
        ext = path.suffix.lower()
        
        if ext == '.xlsx':
            try:
                with zipfile.ZipFile(path, 'r') as zip_file:
                    engine = 'openpyxl'
            except zipfile.BadZipFile:
                engine = 'xlrd'
            except Exception:
                engine = 'openpyxl'
        elif ext == '.xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        df.to_excel(file_path, index=False, engine=engine)


def add_sales_report_period(order_df: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
    """
    为订单明细添加"销售报表账期"列，根据以下规则标记：
    
    1. 识别订单号重复的行：
       - 如果两个订单号一样的订单，合并计算订单金额
       - 结果是0，标记为"全退"
       - 结果不是0，不标记不处理
    
    2. 订单状态为"已取消"并且订单金额为0的订单：
       - 在"销售报表账期"标记为"已取消"
       - 订单金额不是0，不标记不处理
    
    Args:
        order_df: 订单数据DataFrame
        verbose: 是否打印详细日志
    
    Returns:
        添加了"销售报表账期"列的DataFrame
    """
    # 复制DataFrame避免修改原数据
    df = order_df.copy()
    
    # 初始化"销售报表账期"列
    if '销售报表账期' not in df.columns:
        df['销售报表账期'] = None
    else:
        # 清空已有值，重新计算
        df['销售报表账期'] = None
    
    # 确保订单号列为字符串类型
    if '订单号' in df.columns:
        df['订单号'] = df['订单号'].astype(str)
    
    if verbose:
        print("\n=== 开始计算销售报表账期 ===")
        print(f"总行数: {len(df)}")
    
    # 规则1: 识别订单号重复的行，合并计算订单金额
    if '订单号' in df.columns and '订单金额' in df.columns:
        # 将订单金额转为数值类型
        df['_订单金额_numeric'] = pd.to_numeric(df['订单金额'], errors='coerce').fillna(0)
        
        # 按订单号分组，计算金额总和
        order_amount_sum = df.groupby('订单号')['_订单金额_numeric'].sum()
        
        # 找出订单号重复的行（即出现次数 > 1）
        order_counts = df['订单号'].value_counts()
        duplicate_orders = order_counts[order_counts > 1].index.tolist()
        
        if verbose:
            print(f"发现 {len(duplicate_orders)} 个重复订单号")
        
        # 对每个重复订单进行处理
        for order_no in duplicate_orders:
            # 获取该订单号对应的所有行索引
            order_indices = df[df['订单号'] == order_no].index.tolist()
            
            # 计算该订单号的金额总和
            total_amount = order_amount_sum[order_no]
            
            if verbose:
                print(f"  订单号 {order_no}: {len(order_indices)} 行, 金额合计={total_amount}")
            
            # 如果金额合计为0，标记为"全退"
            if total_amount == 0:
                for idx in order_indices:
                    df.at[idx, '销售报表账期'] = '全退'
                if verbose:
                    print(f"    -> 标记为'全退'")
        
        # 删除临时列
        df.drop(columns=['_订单金额_numeric'], inplace=True)
    
    # 规则2: 订单状态为"已取消"且订单金额为0的订单
    if '订单状态' in df.columns and '订单金额' in df.columns:
        # 识别需要标记的行：订单状态包含"取消"且金额为0
        cancel_mask = (
            df['订单状态'].astype(str).str.contains('取消', na=False) &
            (pd.to_numeric(df['订单金额'], errors='coerce') == 0)
        )
        
        # 获取需要标记的行索引
        cancel_indices = df[cancel_mask].index.tolist()
        
        if verbose:
            print(f"发现 {len(cancel_indices)} 个已取消且金额为0的订单")
        
        # 标记为"已取消"
        for idx in cancel_indices:
            # 只有在未标记的情况下才标记（避免覆盖"全退"标记）
            if pd.isna(df.at[idx, '销售报表账期']):
                df.at[idx, '销售报表账期'] = '已取消'
                if verbose:
                    order_no = df.at[idx, '订单号'] if '订单号' in df.columns else 'N/A'
                    print(f"  订单号 {order_no}: 标记为'已取消'")
    
    if verbose:
        # 统计最终标记情况
        marked_count = df['销售报表账期'].notna().sum()
        full_refund_count = (df['销售报表账期'] == '全退').sum()
        cancelled_count = (df['销售报表账期'] == '已取消').sum()
        print(f"\n=== 销售报表账期标记完成 ===")
        print(f"已标记: {marked_count} 行")
        print(f"  - 全退: {full_refund_count} 行")
        print(f"  - 已取消: {cancelled_count} 行")
    
    return df