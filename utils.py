"""
Utility functions for the Excel Merge Tool.
This module contains common functions shared between the interactive and CLI versions.
"""

import pandas as pd
import os
import re
import shutil
from pathlib import Path
from typing import Optional, Any
from datetime import datetime
import logging

# Configure logger for this module
logger = logging.getLogger(__name__)


def auto_backup(file_path: str) -> Path:
    """
    自动备份文件到 backup/ 目录

    Args:
        file_path: 需要备份的文件路径

    Returns:
        Path: 备份文件的路径
    """
    source = Path(file_path)
    if not source.exists():
        return source

    # 创建 backup 目录
    backup_dir = source.parent / "backup"
    backup_dir.mkdir(parents=True, exist_ok=True)

    # 生成带时间戳的备份文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{source.stem}_backup_{timestamp}{source.suffix}"
    backup_path = backup_dir / backup_name

    # 复制文件
    shutil.copy2(source, backup_path)
    logger.info(f"已备份文件到: {backup_path}")

    return backup_path


def extract_p_number(text: Any) -> Optional[str]:
    """
    Extract the part with "P" and following digits from a string
    """
    if pd.isna(text) or text is None:
        return None
    # Convert to string to handle numbers, then search for P pattern
    text_str = str(text)
    match = re.search(r"P\d+", text_str)
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

    if ext == ".csv":
        # For CSV files, try different encodings and parameters if default fails
        # Prioritize gbk encoding for Chinese files
        encodings = ["gbk", "utf-8", "gb2312", "latin-1", "utf-8-sig"]

        df = None

        # First, try to read the file with proper comment handling
        for encoding in encodings:
            try:
                # Read the file as text to check for # comment lines
                with open(file_path, "r", encoding=encoding) as f:
                    lines = f.readlines()

                # Count how many lines start with # at the beginning
                skip_rows = 0
                for line in lines:
                    if line.strip().startswith("#"):
                        skip_rows += 1
                    else:
                        break  # Stop at first line that doesn't start with #

                # Read the CSV file with proper skiprows and encoding
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    skiprows=skip_rows,
                    header=0,  # Use the first non-comment line as header
                    engine="python",  # Use python engine for better encoding handling
                    on_bad_lines="skip",  # Skip bad lines
                )

                # Check if we have at least some expected columns or reasonable data
                if df.shape[0] > 0 and df.shape[1] > 5:
                    break  # We have a valid dataframe

            except UnicodeDecodeError:
                continue
            except pd.errors.ParserError:
                continue
            except Exception as e:
                logger.warning(f"Error reading CSV file with encoding {encoding}: {e}")
                continue

        # If still no success, try with different separators
        if df is None or df.shape[0] == 0:
            for encoding in encodings:
                for sep in [",", ";", "\t"]:
                    try:
                        df = pd.read_csv(
                            file_path,
                            encoding=encoding,
                            sep=sep,
                            engine="python",
                            on_bad_lines="skip",
                        )
                        if df.shape[0] > 0 and df.shape[1] > 5:
                            break
                    except Exception:
                        continue
                if df is not None and df.shape[0] > 0:
                    break

        # If all else fails, try with automatic detection
        if df is None or df.shape[0] == 0:
            df = pd.read_csv(
                file_path,
                encoding="gbk",  # Default to gbk for Chinese files
                engine="python",
                on_bad_lines="skip",
                sep=None,  # Auto-detect separator
                skip_blank_lines=True,
            )

        # Ensure critical columns are treated as strings
        for col in df.columns:
            col_str = str(col)
            if "订单" in col_str or "流水" in col_str:
                df[col] = df[col].astype(str)

        return df
    elif ext in [".xlsx", ".xls"]:
        # For Excel files
        import zipfile

        # Determine engine based on file type
        if ext == ".xlsx":
            try:
                with zipfile.ZipFile(path, "r") as zip_file:
                    # If it's a valid zip file, use openpyxl
                    engine = "openpyxl"
            except zipfile.BadZipFile:
                # If it's not a valid zip file, fall back to xlrd (sometimes older xls files have xlsx extension)
                engine = "xlrd"
            except Exception:
                # For any other error, fall back to openpyxl
                engine = "openpyxl"
        elif ext == ".xls":
            engine = "xlrd"
        else:
            engine = "openpyxl"

        return pd.read_excel(
            file_path,
            dtype={"订单号": str, "商户订单号": str, "商务订单号": str},
            engine=engine,
        )
    else:
        # Default to Excel reading for unknown types (as before)
        try:
            return pd.read_excel(
                file_path,
                dtype={"订单号": str, "商户订单号": str, "商务订单号": str},
                engine="openpyxl",
            )
        except Exception:
            # For CSV files with encoding issues, try different encodings
            encodings = ["utf-8", "gbk", "gb2312", "latin-1"]
            df = None
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if "订单号" in df.columns:
                        df["订单号"] = df["订单号"].astype(str)
                    if "商户订单号" in df.columns:
                        df["商户订单号"] = df["商户订单号"].astype(str)
                    if "商务订单号" in df.columns:
                        df["商务订单号"] = df["商务订单号"].astype(str)
                    return df
                except UnicodeDecodeError:
                    continue  # Try next encoding

            # If all encodings failed, try with utf-8-sig
            df = pd.read_csv(file_path, encoding="utf-8-sig")
            if "订单号" in df.columns:
                df["订单号"] = df["订单号"].astype(str)
            if "商户订单号" in df.columns:
                df["商户订单号"] = df["商户订单号"].astype(str)
            if "商务订单号" in df.columns:
                df["商务订单号"] = df["商务订单号"].astype(str)
            return df


def process_excel_files(
    order_file: str, payment_file: str, verbose: bool = False
) -> pd.DataFrame:
    """
    Process two files (Excel or CSV) according to the specified matching logic.
    Uses more efficient pandas operations instead of nested loops.
    """
    # Read the files using the appropriate method
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)

    # Initialize the '支付手续费' column if it doesn't exist
    if "支付手续费" not in order_df.columns:
        order_df["支付手续费"] = None

    if verbose:
        logger.info("Starting matching process...")
        logger.debug(f"Order file columns: {list(order_df.columns)}")
        logger.debug(f"Payment file columns: {list(payment_df.columns)}")

    # Process each row in the order dataframe
    for idx, order_row in order_df.iterrows():
        if verbose:
            logger.debug(f"\n--- Processing Order Row {idx} ---")
            logger.debug(f"  Full Order Number: {order_row.get('订单号', 'N/A')}")
            logger.debug(
                f"  External Order Number: {order_row.get('外部订单号', 'N/A')}"
            )
            logger.debug(f"  Initial Payment Fee: {order_row.get('支付手续费', 'N/A')}")

        # Get order number (first 20 characters)
        original_order_no = order_row.get("订单号", "")
        order_no = str(original_order_no)[:20] if pd.notna(original_order_no) else ""

        # Don't skip orders with short order numbers, continue for P-number matching
        if pd.isna(original_order_no) or len(str(original_order_no)) < 20:
            if verbose:
                logger.debug(
                    f"Row {idx}: Order number less than 20 characters: {original_order_no}, continuing for P-number matching"
                )
        external_order_no = order_row.get("外部订单号", None)

        if verbose:
            logger.debug(f"  Truncated Order Number (first 20 chars): {order_no}")

        # Determine if it's a regular order, refund order, or skip
        order_amount_raw = order_row.get("订单金额", 0)
        if verbose:
            logger.debug(f"  Raw Order Amount: {order_amount_raw}")

        if pd.isna(order_amount_raw):
            order_amount = 0  # Treat NaN as 0
            if verbose:
                logger.debug("  Order amount is NaN, setting to 0")
        else:
            # Ensure it's a numeric value to avoid issues with string values
            try:
                order_amount = float(order_amount_raw)
                if verbose:
                    logger.debug(f"  Converted Order Amount: {order_amount}")
            except (ValueError, TypeError):
                order_amount = 0  # Default to 0 if conversion fails
                if verbose:
                    logger.debug(
                        f"  Failed to convert amount '{order_amount_raw}' to float, setting to 0"
                    )

        # Updated logic: positive amounts > 0 = regular order, negative amounts < 0 = refund, amount = 0 = set 支付手续费 to 0
        if order_amount > 0:
            is_regular_order = True
            order_type = "正单(Regular)"
        elif order_amount < 0:
            is_regular_order = False
            order_type = "退单(Refund)"
        else:  # order_amount == 0
            if verbose:
                logger.debug(f"Row {idx}: Order amount is 0, setting 支付手续费 to 0")
            order_df.at[idx, "支付手续费"] = 0.0
            continue  # Skip further processing for this row but set the fee to 0

        if verbose:
            logger.debug(
                f"Row {idx}: Processing - Order No: {order_no}, External Order: {external_order_no}, Amount: {order_amount} ({order_type})"
            )

        # Find matching records in payment dataframe using vectorized operations where possible
        matching_payments = []

        # First, try to find exact matches by truncated order number
        # Check if '商户订单号' column exists, if not, try other possible column names
        business_order_col = None
        for col in payment_df.columns:
            col_str = str(col)
            if "商户" in col_str and "订单" in col_str:
                business_order_col = col
                break

        if not business_order_col:
            # If no column with '商户订单号' found, try other possible column names
            for col in payment_df.columns:
                if "订单" in str(col):
                    business_order_col = col
                    break

        if business_order_col:
            business_order_numbers = payment_df[business_order_col].astype(str)
            exact_matches = business_order_numbers.str[:20] == order_no
            exact_match_rows = payment_df[exact_matches]

            if verbose and len(exact_match_rows) > 0:
                logger.debug(
                    f"  Found {len(exact_match_rows)} exact matches for order {order_no}"
                )
        else:
            # If no column with '订单' found, skip exact match check
            exact_match_rows = payment_df.head(0)  # Empty dataframe
            if verbose:
                logger.debug(f"  No business order column found in payment file")

        # For non-exact matches, check P-number and hyphen logic
        if len(exact_match_rows) > 0:
            # Handle exact matches
            for p_idx, payment_row in exact_match_rows.iterrows():
                # Check business type before adding to matching payments
                business_type = payment_row.get("业务类型", "")
                business_type_str = str(business_type).strip()

                # Check if this payment has the right business type for the order
                is_payment_charge = (
                    "收费" in business_type_str or "服务费" in business_type_str
                )
                is_payment_refund = (
                    "退费" in business_type_str or "退款" in business_type_str
                )

                # For regular orders, we only want charge type payments that have non-zero expenditure
                if is_regular_order and is_payment_charge:
                    expenditure = payment_row.get("支出金额（-元）", 0)
                    try:
                        expenditure_val = float(expenditure)
                        if expenditure_val != 0:
                            # Regular order (positive amount) matches with charge type and non-zero expenditure
                            matching_payments.append(payment_row)
                            if verbose:
                                logger.debug(
                                    f"    - Added exact match at payment row {p_idx} (charge type with non-zero expenditure)"
                                )
                    except (ValueError, TypeError):
                        if verbose:
                            logger.debug(
                                f"    - Skipped exact match at payment row {p_idx} (invalid expenditure value)"
                            )
                # For refund orders, we only want refund type payments that have non-zero income
                elif not is_regular_order and is_payment_refund:
                    income = payment_row.get("收入金额（+元）", 0)
                    try:
                        income_val = float(income)
                        if income_val != 0:
                            # Refund order (negative amount) matches with refund type and non-zero income
                            matching_payments.append(payment_row)
                            if verbose:
                                logger.debug(
                                    f"    - Added exact match at payment row {p_idx} (refund type with non-zero income)"
                                )
                    except (ValueError, TypeError):
                        if verbose:
                            logger.debug(
                                f"    - Skipped exact match at payment row {p_idx} (invalid income value)"
                            )
                else:
                    # Skip payments with incorrect business type
                    if verbose:
                        logger.debug(
                            f"    - Skipped exact match at payment row {p_idx} (incorrect business type: {business_type_str})"
                        )
        
        if len(matching_payments) == 0:
            # If no exact matches were found, check P-number and hyphen logic
            for p_idx, payment_row in payment_df.iterrows():
                # Try P-number match
                product_name = payment_row.get("商品名称", "")
                p_number_match = False

                # Enhanced P-number matching
                p_number_match = match_orders_by_p_number(
                    external_order_no, product_name
                )
                if verbose and p_number_match:
                    external_p = extract_p_number(external_order_no)
                    product_p = extract_p_number(product_name)
                    logger.debug(f"      P-number match: {external_p} == {product_p}")

                # Enhanced hyphen match
                hyphen_match = False
                if pd.notna(product_name) and pd.notna(external_order_no):
                    external_str = str(external_order_no).strip()
                    product_str = str(product_name).strip()

                    if "-" in product_str:
                        # The logic is to match any segment of external_order_no (split by '-')
                        # with the segment *after the last hyphen* of product_name
                        last_part = product_str.rsplit("-", 1)[-1]
                        external_parts = external_str.split("-") if "-" in external_str else [external_str]
                        if last_part and last_part in external_parts:
                            hyphen_match = True
                            if verbose:
                                logger.debug(
                                    f"      Hyphen match: '{last_part}' in {external_parts}"
                                )

                # Check business type
                business_type = payment_row.get("业务类型", "")
                business_type_str = str(business_type).strip()

                # Check if business type matches expected values
                is_charge = False
                is_refund = False

                if "收费" in business_type_str:
                    is_charge = True
                elif "退费" in business_type_str:
                    is_refund = True
                elif "退款" in business_type_str:
                    is_refund = True
                elif "服务费" in business_type_str:
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
                    logger.debug(
                        f"      Checking payment row {p_idx}: type='{business_type_str}', charge={is_charge}, refund={is_refund}, correct={business_type_correct}"
                    )

                # If we have a match based on P-number or hyphen AND correct business type, add to matching payments
                if (p_number_match or hyphen_match):
                    # For non-exact matches, we also need to verify the business type before confirming the match.
                    if business_type_correct:
                        matching_payments.append(payment_row)
                        if verbose:
                            logger.debug(f"    - Match confirmed at payment row {p_idx}")
                        # Since we found a valid match, we can break the inner loop to prevent subsequent rows from overwriting it.
                        break
                    elif verbose:
                        logger.debug(
                            f"    - Skipped match at payment row {p_idx} (incorrect business type: {business_type_str})"
                        )

        # If matches found, get the appropriate amount and update '支付手续费'
        if matching_payments:
            # Find the appropriate amount columns
            income_col = "收入金额（+元）"
            expenditure_col = "支出金额（-元）"

            if verbose:
                logger.debug(
                    f"  Using income column: {income_col}, expenditure column: {expenditure_col}"
                )

            updated = False

            # For all matching payments, find the one with correct business type and amount
            for payment in matching_payments:
                business_type = payment.get("业务类型", "")
                business_type_str = str(business_type).strip()

                # Check if this payment has the right business type for the order
                is_payment_charge = (
                    "收费" in business_type_str or "服务费" in business_type_str
                )
                is_payment_refund = (
                    "退费" in business_type_str or "退款" in business_type_str
                )

                if is_regular_order:
                    # For regular order: 支付手续费 = 支出金额（-元）的值且需要转成负值，代表需要付出去的钱
                    if is_payment_charge:
                        expenditure = payment.get(expenditure_col, 0)
                        try:
                            expenditure_val = float(expenditure)
                            if expenditure_val != 0:
                                # 支出金额已经是负数，我们需要保持其负值状态
                                order_df.at[idx, "支付手续费"] = expenditure_val
                                if verbose:
                                    logger.debug(
                                        f"  - Updated 支付手续费 for regular order: {expenditure_val}"
                                    )
                                updated = True
                                break
                        except (ValueError, TypeError):
                            if verbose:
                                logger.debug(
                                    f"  - Invalid expenditure value: {expenditure}"
                                )
                else:
                    # For refund order: 支付手续费 = 收入金额（+元）的值且应该为正数，代表退单后需要收回来的钱
                    if is_payment_refund:
                        income = payment.get(income_col, 0)
                        try:
                            income_val = float(income)
                            if income_val != 0:
                                # 收入金额已经是正数，我们需要保持其正数状态
                                order_df.at[idx, "支付手续费"] = income_val
                                if verbose:
                                    logger.debug(
                                        f"  - Updated 支付手续费 for refund order: {income_val}"
                                    )
                                updated = True
                                break
                        except (ValueError, TypeError):
                            if verbose:
                                logger.debug(f"  - Invalid income value: {income}")

            # If no valid payment found, do not update the payment fee
            if not updated and verbose:
                logger.debug(f"  - No valid payment amount found for this order")
        else:
            if verbose:
                logger.debug(f"  - No matches found for this order")

    if verbose:
        logger.info("Matching process completed.")

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
    if original_file_extension.lower() == ".csv":
        df.to_csv(file_path, index=False, encoding="utf-8-sig")
    else:
        # For Excel files, determine the appropriate engine
        path = Path(file_path)
        ext = path.suffix.lower()

        if ext == ".xlsx":
            try:
                with zipfile.ZipFile(path, "r") as zip_file:
                    engine = "openpyxl"
            except zipfile.BadZipFile:
                engine = "xlrd"
            except Exception:
                engine = "openpyxl"
        elif ext == ".xls":
            engine = "xlrd"
        else:
            engine = "openpyxl"

        df.to_excel(file_path, index=False, engine=engine)


def add_sales_report_period(
    order_df: pd.DataFrame, verbose: bool = False
) -> pd.DataFrame:
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
    if "销售报表账期" not in df.columns:
        df["销售报表账期"] = None
    else:
        # 清空已有值，重新计算
        df["销售报表账期"] = None

    # 确保订单号列为字符串类型
    if "订单号" in df.columns:
        df["订单号"] = df["订单号"].astype(str)

    if verbose:
        logger.info("\n=== 开始计算销售报表账期 ===")
        logger.info(f"总行数: {len(df)}")

    # 规则1: 识别订单号重复的行，合并计算订单金额
    if "订单号" in df.columns and "订单金额" in df.columns:
        # 将订单金额转为数值类型
        df["_订单金额_numeric"] = pd.to_numeric(df["订单金额"], errors="coerce").fillna(
            0
        )

        # 按订单号分组，计算金额总和
        order_amount_sum = df.groupby("订单号")["_订单金额_numeric"].sum()

        # 找出订单号重复的行（即出现次数 > 1）
        order_counts = df["订单号"].value_counts()
        duplicate_orders = order_counts[order_counts > 1].index.tolist()

        if verbose:
            logger.info(f"发现 {len(duplicate_orders)} 个重复订单号")

        # 对每个重复订单进行处理
        for order_no in duplicate_orders:
            # 获取该订单号对应的所有行索引
            order_indices = df[df["订单号"] == order_no].index.tolist()

            # 计算该订单号的金额总和
            total_amount = order_amount_sum[order_no]

            if verbose:
                logger.debug(
                    f"  订单号 {order_no}: {len(order_indices)} 行, 金额合计={total_amount}"
                )

            # 如果金额合计为0，标记为"全退"
            if total_amount == 0:
                for idx in order_indices:
                    df.at[idx, "销售报表账期"] = "全退"
                if verbose:
                    logger.debug(f"    -> 标记为'全退'")

        # 删除临时列
        df.drop(columns=["_订单金额_numeric"], inplace=True)

    # 规则2: 订单状态为"已取消"且订单金额为0的订单
    if "订单状态" in df.columns and "订单金额" in df.columns:
        # 识别需要标记的行：订单状态包含"取消"且金额为0
        cancel_mask = df["订单状态"].astype(str).str.contains("取消", na=False) & (
            pd.to_numeric(df["订单金额"], errors="coerce") == 0
        )

        # 获取需要标记的行索引
        cancel_indices = df[cancel_mask].index.tolist()

        if verbose:
            logger.info(f"发现 {len(cancel_indices)} 个已取消且金额为0的订单")

        # 标记为"已取消"
        for idx in cancel_indices:
            # 只有在未标记的情况下才标记（避免覆盖"全退"标记）
            if pd.isna(df.at[idx, "销售报表账期"]):
                df.at[idx, "销售报表账期"] = "已取消"
                if verbose:
                    order_no = df.at[idx, "订单号"] if "订单号" in df.columns else "N/A"
                    logger.debug(f"  订单号 {order_no}: 标记为'已取消'")

    if verbose:
        # 统计最终标记情况
        marked_count = df["销售报表账期"].notna().sum()
        full_refund_count = (df["销售报表账期"] == "全退").sum()
        cancelled_count = (df["销售报表账期"] == "已取消").sum()
        logger.info(f"\n=== 销售报表账期标记完成 ===")
        logger.info(f"已标记: {marked_count} 行")
        logger.info(f"  - 全退: {full_refund_count} 行")
        logger.info(f"  - 已取消: {cancelled_count} 行")

    return df


def parse_date(date_val: Any) -> Optional[pd.Timestamp]:
    """
    解析日期值，支持多种格式

    Args:
        date_val: 日期值，可以是字符串、datetime、pandas Timestamp等

    Returns:
        pandas Timestamp 或 None（如果解析失败）
    """
    if pd.isna(date_val) or date_val is None:
        return None

    if isinstance(date_val, pd.Timestamp):
        return date_val

    if isinstance(date_val, (datetime,)):
        try:
            return pd.Timestamp(date_val)
        except Exception:
            pass

    text = str(date_val).strip()

    try:
        return pd.to_datetime(text)
    except Exception:
        pass

    import re

    chinese_match = re.match(r"(\d{4})[年](\d{1,2})[月](\d{0,2})", text)
    if chinese_match:
        year, month = int(chinese_match.group(1)), int(chinese_match.group(2))
        return pd.Timestamp(year=year, month=month, day=1)

    return None


def get_year_month(date_val: Any) -> Optional[str]:
    """
    从日期值获取年月字符串，格式为 YYYYMM

    Args:
        date_val: 日期值

    Returns:
        年月字符串（如 "202602"）或 None
    """
    parsed = parse_date(date_val)
    if parsed is None:
        return None

    return parsed.strftime("%Y%m")


def filter_unmarked_and_generate_report(
    order_df: pd.DataFrame,
    target_month: str,
    verbose: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    第二阶段功能：筛选未标记数据并在订单 DataFrame 上回填账期标记。

    本工作流不写出任何文件；返回的 DataFrame 仅供调用方决定如何持久化。

    流程：
    1. 过滤掉已标记的数据（"销售报表账期"列不为空的行）
    2. 筛选出"出发日期"指定月份（例如2026年2月）的数据
    3. 往前查一年（例如2026年2月往前查一年是2025年2月到2026年2月）
    4. 从所有未标记的数据中，找出"出发日期"在这一年内范围的所有数据
    5. 将筛选出的行作为内存中的"月报 DataFrame"返回（不落盘）
    6. 同时，将被复制的原Excel表格中这些被复制的数据行，
       在"销售报表账期"列填上"销售报表YYYYMM"

    Args:
        order_df: 订单数据DataFrame（已执行第一阶段标记）
        target_month: 目标月份，格式为 YYYYMM（如 "202602"）
        verbose: 是否打印详细日志

    Returns:
        tuple: (更新后的原DataFrame, 筛选得到的月报DataFrame; 均为内存对象)
    """
    # 复制DataFrame避免修改原数据
    df = order_df.copy()

    if verbose:
        logger.info("\n" + "=" * 60)
        logger.info("开始第二阶段：筛选未标记数据并生成新文档")
        logger.info("=" * 60)
        logger.info(f"目标月份: {target_month}")
        logger.info(f"原始数据总行数: {len(df)}")

    # 确保"销售报表账期"列存在
    if "销售报表账期" not in df.columns:
        df["销售报表账期"] = None

    # 步骤1: 过滤掉已标记的数据（"销售报表账期"列不为空的行）
    unmarked_mask = df["销售报表账期"].isna()
    unmarked_df = df[unmarked_mask].copy()

    if verbose:
        marked_count = (~unmarked_mask).sum()
        logger.info(f"\n步骤1: 过滤已标记数据")
        logger.info(f"  已标记行数: {marked_count}")
        logger.info(f"  未标记行数: {len(unmarked_df)}")

    # 步骤2 & 3 & 4: 筛选出发日期在指定月份往前一年范围内的数据
    # 支持"出发日期"和"出行日期"两种列名
    date_col = None
    for col in ["出发日期", "出行日期"]:
        if col in unmarked_df.columns:
            date_col = col
            break

    if date_col is None:
        if verbose:
            logger.warning("\n警告: 数据中没有'出发日期'或'出行日期'列，跳过日期筛选")
        # 如果没有日期列，返回空的筛选结果
        return df, pd.DataFrame()
    
    if verbose:
        logger.info(f"\n使用日期列: {date_col}")
    # 解析目标月份
    try:
        target_year = int(target_month[:4])
        target_month_num = int(target_month[4:6])
    except (ValueError, IndexError):
        if verbose:
            logger.warning(f"无效的目标月份格式: {target_month}，跳过日期筛选")
        return df, pd.DataFrame()

    # Date window per sales-report spec: target month ± 1 year
    # Start: first day of (target_year - 1, target_month_num)
    # End:   last day of  (target_year + 1, target_month_num)
    start_date = pd.Timestamp(year=target_year - 1, month=target_month_num, day=1)
    end_date = pd.Timestamp(year=target_year + 1, month=target_month_num, day=1) + pd.offsets.MonthEnd(0)


    if verbose:
        logger.info(f"\n步骤2 & 3 & 4: 筛选出行日期")
        logger.info(f"  目标月份: {target_month}")
        logger.info(f"  往前查一年范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")

    # 解析每行的出发日期，筛选在范围内且未标记的数据
    unmarked_df['_travel_date'] = pd.to_datetime(unmarked_df[date_col], errors='coerce')
    
    date_filter_mask = (unmarked_df['_travel_date'] >= start_date) & (unmarked_df['_travel_date'] <= end_date)
    
    filtered_df = unmarked_df[date_filter_mask].copy()
    filtered_df.drop(columns=['_travel_date'], inplace=True)


    if verbose:
        logger.info(f"  符合条件的数据行数: {len(filtered_df)}")

    # 步骤5: 收集筛选结果为内存中的月报 DataFrame（不落盘）
    new_report_df = pd.DataFrame()
    if len(filtered_df) > 0:
        new_report_df = filtered_df.copy()

        if verbose:
            logger.info(f"\n步骤5: 收集月报 DataFrame")
            logger.info(f"  行数: {len(new_report_df)}")
            logger.info(f"  注意: 工作流不写出报表文件；如需持久化请由调用方处理")

    # 步骤6: 在原Excel中标记被复制的数据行
    if len(filtered_df) > 0:
        mark_value = f"销售报表{target_month}"

        for idx in filtered_df.index:
            df.at[idx, "销售报表账期"] = mark_value

        if verbose:
            logger.info(f"\n步骤6: 在原Excel中标记被复制的数据")
            logger.info(f"  标记值: {mark_value}")
            logger.info(f"  标记行数: {len(filtered_df)}")

    if verbose:
        logger.info("\n" + "=" * 60)
        logger.info("第二阶段处理完成")
        logger.info("=" * 60)

    return df, new_report_df


def process_sales_report_workflow(
    order_file: str,
    payment_file: str,
    target_month: str,
    verbose: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    完整的销售报表工作流：处理两个文件并计算月报 DataFrame。

    本工作流不写出任何文件；返回的 DataFrame 仅供调用方决定如何持久化
    （CLI 与交互模式仅就地写回订单文件，不产生独立报表；HTTP API 自行
    决定是否将月报 DataFrame 落盘以服务下载）。

    流程：
    1. 读取并处理订单文件和支付文件（匹配支付手续费）
    2. 执行第一阶段：添加"销售报表账期"标记（全退、已取消）
    3. 执行第二阶段：筛选未标记数据并回填账期标记（不落盘）

    Args:
        order_file: 订单文件路径
        payment_file: 支付文件路径
        target_month: 目标月份（格式 YYYYMM）
        verbose: 是否打印详细日志

    Returns:
        tuple: (更新后的订单DataFrame, 月报DataFrame; 均为内存对象)
    """
    if verbose:
        logger.info("=" * 60)
        logger.info("启动销售报表工作流")
        logger.info("=" * 60)
        logger.info(f"订单文件: {order_file}")
        logger.info(f"支付文件: {payment_file}")
        logger.info(f"目标月份: {target_month}")

    # 步骤1: 处理订单文件和支付文件
    result_df = process_excel_files(order_file, payment_file, verbose=verbose)

    # 步骤2 & 3: 第二阶段处理（不落盘）
    updated_df, report_df = filter_unmarked_and_generate_report(
        result_df, target_month, verbose=verbose
    )

    return updated_df, report_df
