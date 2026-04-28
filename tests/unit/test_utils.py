import pandas as pd
from pandas.testing import assert_frame_equal
import pytest
from pathlib import Path
from utils import (
    read_file_with_appropriate_method,
    process_excel_files,
    add_sales_report_period,
    filter_unmarked_and_generate_report,
)

def test_read_file_with_appropriate_method(sample_data_dir):
    """
    Task 2.5: Verify that files with different extensions are read correctly.
    """
    excel_path = sample_data_dir / "orders.xlsx"
    csv_path = sample_data_dir / "orders.csv"

    # Test reading .xlsx
    df_excel = read_file_with_appropriate_method(excel_path)
    assert isinstance(df_excel, pd.DataFrame)
    assert not df_excel.empty
    assert "订单号" in df_excel.columns

    # Test reading .csv
    df_csv = read_file_with_appropriate_method(csv_path)
    assert isinstance(df_csv, pd.DataFrame)
    assert not df_csv.empty
    assert "订单号" in df_csv.columns

    # Check that content is roughly the same (ignoring type differences from read)
    assert df_excel.shape == df_csv.shape

def test_process_excel_files_matching_logic(sample_data_dir):
    """
    Task 2.2: Test the core matching logic in process_excel_files.
    Covers exact, P-number, and hyphen matching, plus business type validation.
    """
    order_path = sample_data_dir / "orders.xlsx"
    payment_path = sample_data_dir / "payments.csv"

    result_df = process_excel_files(order_path, payment_path, verbose=False)

    # Expected fees after matching
    expected_fees = {
        "EXACT_MATCH_ORDER_12345": -10.0,
        "EXACT_MATCH_REFUND_67890": 5.0,
        "ORDER_WITH_PNUM_ABCDE": -20.0,
        "ORDER_WITH_HYPHEN_FGHIJ": -30.0,
        "ZERO_AMOUNT_ORDER_444": 0.0,
    }

    for order_no, fee in expected_fees.items():
        matched_fee = result_df.loc[result_df["订单号"] == order_no, "支付手续费"].iloc[0]
        assert matched_fee == fee, f"Mismatch for order {order_no}"

    # Verify that an unmatched order has no fee
    unmatched_fee = result_df.loc[result_df["订单号"] == "UNMATCHED_ORDER_333", "支付手续费"].iloc[0]
    assert pd.isna(unmatched_fee), "Unmatched order should have NaN fee"

def test_add_sales_report_period_marking(sample_data_dir):
    """
    Task 2.3: Test the marking logic for "全退" (Full Refund) and "已取消" (Cancelled).
    """
    order_df = read_file_with_appropriate_method(sample_data_dir / "orders.xlsx")
    result_df = add_sales_report_period(order_df.copy(), verbose=False)

    # Check "全退"
    full_refund_marks = result_df[result_df["订单号"] == "FULL_REFUND_ORDER_111"]["销售报表账期"]
    assert all(mark == "全退" for mark in full_refund_marks), 'Full refund not marked correctly'

    # Check "已取消"
    cancelled_mark = result_df.loc[result_df["订单号"] == "CANCELLED_ORDER_222", "销售报表账期"].iloc[0]
    assert cancelled_mark == "已取消", "Cancelled order not marked correctly"

    # Check that other orders are not marked
    other_mark = result_df.loc[result_df["订单号"] == "EXACT_MATCH_ORDER_12345", "销售报表账期"].iloc[0]
    assert pd.isna(other_mark), "Other orders should not be marked"


def test_filter_unmarked_and_generate_report(sample_data_dir, tmp_path):
    """
    Task 2.4: Test the filtering logic for the monthly sales report.
    Ensures only relevant, unmarked rows within the date window are included.
    """
    order_df = read_file_with_appropriate_method(sample_data_dir / "orders.xlsx")
    
    # Pre-mark the dataframe to simulate a real workflow
    marked_df = add_sales_report_period(order_df.copy(), verbose=False)

    updated_df, report_df_generated = filter_unmarked_and_generate_report(
        order_df=marked_df,
        target_month="202603",
        output_dir=str(tmp_path),
        verbose=False
    )
    
    # After filtering, a report file should be generated in the temp directory
    report_files = list(tmp_path.glob("report_*.xlsx"))
    assert len(report_files) > 0, "Report file was not generated"
    report_path = report_files[0]
    
    report_df = pd.read_excel(report_path)

    # Only one row should be in the report: DATE_FILTER_THIS_MONTH
    # DATE_FILTER_LAST_YEAR is outside the month but inside the 1-year window
    # DATE_FILTER_OUT_OF_RANGE is outside the 1-year window
    # FULL_REFUND_ORDER and CANCELLED_ORDER are marked and should be excluded
    expected_rows = 2
    assert report_df.shape[0] == expected_rows

    included_orders = report_df["订单号"].tolist()
    assert "DATE_FILTER_THIS_MONTH" in included_orders
    assert "DATE_FILTER_LAST_YEAR" in included_orders
    assert "DATE_FILTER_OUT_OF_RANGE" not in included_orders
    assert "FULL_REFUND_ORDER_111" not in included_orders
    assert "CANCELLED_ORDER_222" not in included_orders

    # Check that the updated_df has the correct markings
    mark_value = f"销售报表202603"
    assert updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_THIS_MONTH', '销售报表账期'].iloc[0] == mark_value
    assert updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_LAST_YEAR', '销售报表账期'].iloc[0] == mark_value
    assert pd.isna(updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_OUT_OF_RANGE', '销售报表账期'].iloc[0])
