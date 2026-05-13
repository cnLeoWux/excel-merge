"""Sales report workflow for Excel Merge Tool."""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Optional, Tuple

import pandas as pd
from pandas.errors import OutOfBoundsDatetime

logger = logging.getLogger(__name__)


def add_sales_report_period(order_df: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
    df = order_df.copy()
    df["销售报表账期"] = None
    if "订单号" in df.columns:
        df["订单号"] = df["订单号"].fillna("").astype(str).str.strip('="\t ')
    if "订单号" in df.columns and "订单金额" in df.columns:
        df["_订单金额_numeric"] = pd.to_numeric(df["订单金额"], errors="coerce").fillna(0)
        order_amount_sum = df.groupby("订单号")["_订单金额_numeric"].sum()
        order_counts = df["订单号"].value_counts()
        for order_no in order_counts[order_counts > 1].index.tolist():
            if order_amount_sum[order_no] == 0:
                for idx in df[df["订单号"] == order_no].index.tolist():
                    df.at[idx, "销售报表账期"] = "全退"
        df.drop(columns=["_订单金额_numeric"], inplace=True)
    if "订单状态" in df.columns and "订单金额" in df.columns:
        cancel_mask = df["订单状态"].fillna("").astype(str).str.contains("取消", na=False) & (pd.to_numeric(df["订单金额"], errors="coerce") == 0)
        for idx in df[cancel_mask].index.tolist():
            if pd.isna(df.at[idx, "销售报表账期"]):
                df.at[idx, "销售报表账期"] = "已取消"
    return df


def parse_date(date_val: Any) -> Optional[pd.Timestamp]:
    if pd.isna(date_val) or date_val is None:
        return None
    if isinstance(date_val, pd.Timestamp):
        return date_val if not pd.isna(date_val) else None
    if isinstance(date_val, datetime):
        try:
            result = pd.Timestamp(date_val)
            return None if pd.isna(result) else result
        except (ValueError, TypeError, OutOfBoundsDatetime):
            return None
    text = str(date_val).strip()
    try:
        result = pd.to_datetime(text)
        return None if pd.isna(result) else result
    except (ValueError, TypeError, OutOfBoundsDatetime):
        pass
    import re
    chinese_match = re.match(r"(\d{4})[年](\d{1,2})[月](\d{0,2})", text)
    if chinese_match:
        return pd.Timestamp(year=int(chinese_match.group(1)), month=int(chinese_match.group(2)), day=1)
    return None


def get_year_month(date_val: Any) -> Optional[str]:
    parsed = parse_date(date_val)
    return parsed.strftime("%Y%m") if parsed is not None else None


def filter_unmarked_and_generate_report(
    order_df: pd.DataFrame, target_month: str, verbose: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = order_df.copy()
    if "销售报表账期" not in df.columns:
        df["销售报表账期"] = None
    unmarked_df = df[df["销售报表账期"].isna()].copy()
    date_col = None
    for col in ["出发日期", "出行日期"]:
        if col in unmarked_df.columns:
            date_col = col
            break
    if date_col is None:
        return df, pd.DataFrame()
    try:
        target_year = int(target_month[:4])
        target_month_num = int(target_month[4:6])
    except (ValueError, IndexError):
        return df, pd.DataFrame()
    start_date = pd.Timestamp(year=target_year - 1, month=target_month_num, day=1)
    end_date = pd.Timestamp(year=target_year + 1, month=target_month_num, day=1) + pd.offsets.MonthEnd(0)
    unmarked_df["_travel_date"] = pd.to_datetime(unmarked_df[date_col], errors="coerce")
    filtered_df = unmarked_df[(unmarked_df["_travel_date"] >= start_date) & (unmarked_df["_travel_date"] <= end_date)].copy()
    filtered_df.drop(columns=["_travel_date"], inplace=True)
    if len(filtered_df) > 0:
        mark_value = f"销售报表{target_month}"
        for idx in filtered_df.index:
            df.at[idx, "销售报表账期"] = mark_value
    return df, filtered_df.copy() if len(filtered_df) > 0 else pd.DataFrame()


def process_sales_report_workflow(
    order_file: str, payment_file: str, target_month: str, verbose: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    from matching import process_excel_files

    result_df = process_excel_files(order_file, payment_file, verbose=verbose)
    return filter_unmarked_and_generate_report(result_df, target_month, verbose=verbose)
