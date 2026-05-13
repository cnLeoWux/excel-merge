"""Matching workflow for Excel Merge Tool."""

from __future__ import annotations

import logging
import re
from typing import Any, Optional

import pandas as pd

from file_io import read_file_with_appropriate_method

logger = logging.getLogger(__name__)


def extract_p_number(text: Any) -> Optional[str]:
    if pd.isna(text) or text is None:
        return None
    match = re.search(r"P\d+", str(text))
    return match.group() if match else None


def match_orders_by_p_number(external_order_no: Any, product_name: Any) -> bool:
    external_p = extract_p_number(external_order_no)
    product_p = extract_p_number(product_name)
    return bool(external_p and product_p and external_p == product_p)


def _detect_business_order_column(payment_df: pd.DataFrame):
    for col in payment_df.columns:
        if "商户" in str(col) and "订单" in str(col):
            return col
    for col in payment_df.columns:
        if "订单" in str(col):
            return col
    return None


def _classify_order_amount(order_amount_raw: Any) -> float:
    if order_amount_raw is None or pd.isna(order_amount_raw):
        return 0.0
    try:
        return float(order_amount_raw)
    except (ValueError, TypeError):
        return 0.0


def _is_business_type_compatible(order_amount: float, business_type: Any) -> bool:
    bt = str(business_type).strip()
    if order_amount > 0:
        return "收费" in bt or "服务费" in bt
    if order_amount < 0:
        return "退费" in bt or "退款" in bt
    return False


def _extract_payment_fee(payment_row: pd.Series, is_regular_order: bool):
    if is_regular_order:
        try:
            raw = payment_row.get("支出金额（-元）", 0)
            val = float(0 if raw is None or pd.isna(raw) else raw)
            return val if val != 0 else None
        except (ValueError, TypeError):
            return None
    try:
        raw = payment_row.get("收入金额（+元）", 0)
        val = float(0 if raw is None or pd.isna(raw) else raw)
        return val if val != 0 else None
    except (ValueError, TypeError):
        return None


def _matches_exact_order(order_no: str, payment_order_no: Any) -> bool:
    return str(payment_order_no).strip('="\t ')[:20] == order_no


def _matches_hyphen_fallback(external_order_no: Any, product_name: Any) -> bool:
    if pd.isna(product_name) or pd.isna(external_order_no):
        return False
    external_str = str(external_order_no).strip()
    product_str = str(product_name).strip()
    if "-" not in product_str:
        return False
    last_part = product_str.rsplit("-", 1)[-1]
    external_parts = external_str.split("-") if "-" in external_str else [external_str]
    return bool(last_part and last_part in external_parts)


def _find_exact_match(order_no: str, payment_df: pd.DataFrame, business_order_col):
    if business_order_col is None:
        return payment_df.head(0)
    mask = payment_df[business_order_col].fillna("").astype(str).str.strip('="\t ').str[:20] == order_no
    return payment_df[mask]


def _find_fallback_match(order_amount: float, external_order_no: Any, payment_df: pd.DataFrame, is_regular_order: bool):
    for _, payment_row in payment_df.iterrows():
        p_number_match = match_orders_by_p_number(external_order_no, payment_row.get("商品名称", ""))
        hyphen_match = _matches_hyphen_fallback(external_order_no, payment_row.get("商品名称", ""))
        if not (p_number_match or hyphen_match):
            continue
        if _is_business_type_compatible(order_amount, payment_row.get("业务类型", "")):
            return payment_row
    return None


def process_excel_files(order_file: str, payment_file: str, verbose: bool = False) -> pd.DataFrame:
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)
    if "支付手续费" not in order_df.columns:
        order_df["支付手续费"] = None
    business_order_col = _detect_business_order_column(payment_df)

    for idx, order_row in order_df.iterrows():
        original_order_no = order_row.get("订单号", "")
        order_no = "" if original_order_no is None or pd.isna(original_order_no) else str(original_order_no)[:20]
        order_amount = _classify_order_amount(order_row.get("订单金额", 0))
        external_order_no = order_row.get("外部订单号", None)
        if order_amount == 0:
            order_df.at[idx, "支付手续费"] = 0.0
            continue
        is_regular_order = order_amount > 0
        exact_match_rows = _find_exact_match(order_no, payment_df, business_order_col)
        matching_fee = None
        for _, payment_row in exact_match_rows.iterrows():
            if _is_business_type_compatible(order_amount, payment_row.get("业务类型", "")):
                matching_fee = _extract_payment_fee(payment_row, is_regular_order)
                if matching_fee is not None:
                    break
        if matching_fee is None:
            fallback_row = _find_fallback_match(order_amount, external_order_no, payment_df, is_regular_order)
            if fallback_row is not None:
                matching_fee = _extract_payment_fee(fallback_row, is_regular_order)
        if matching_fee is not None:
            order_df.at[idx, "支付手续费"] = matching_fee

    from sales_report import add_sales_report_period

    return add_sales_report_period(order_df, verbose=verbose)
