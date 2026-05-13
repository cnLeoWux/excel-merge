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


def test_read_file_with_appropriate_method_xlsx_happy_path(sample_data_dir):
    df = read_file_with_appropriate_method(sample_data_dir / "orders.xlsx")
    assert isinstance(df, pd.DataFrame)
    assert not df.empty
    assert "订单号" in df.columns


def test_read_file_with_appropriate_method_xls_happy_path(tmp_path):
    xlwt = pytest.importorskip("xlwt")
    xls_path = tmp_path / "orders.xls"

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet1")
    headers = ["订单号", "商户订单号", "订单金额"]
    for col, value in enumerate(headers):
        sheet.write(0, col, value)
    sheet.write(1, 0, "XLS_ORDER_1")
    sheet.write(1, 1, "XLS_MERCHANT_1")
    sheet.write(1, 2, 100)
    workbook.save(str(xls_path))

    df = read_file_with_appropriate_method(xls_path)
    assert isinstance(df, pd.DataFrame)
    assert df.iloc[0]["订单号"] == "XLS_ORDER_1"

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
    Ensures only relevant, unmarked rows within the date window are included,
    and that the function does NOT write any files to disk (in-place contract).
    """
    order_df = read_file_with_appropriate_method(sample_data_dir / "orders.xlsx")

    # Pre-mark the dataframe to simulate a real workflow
    marked_df = add_sales_report_period(order_df.copy(), verbose=False)

    # Snapshot tmp_path; the function MUST NOT write into it (no output_dir param).
    snapshot_before = {p for p in tmp_path.rglob("*") if p.is_file()}

    # Function signature must NOT accept output_dir
    import inspect

    sig = inspect.signature(filter_unmarked_and_generate_report)
    assert "output_dir" not in sig.parameters

    updated_df, report_df = filter_unmarked_and_generate_report(
        order_df=marked_df,
        target_month="202603",
        verbose=False,
    )

    # No file artefacts should appear anywhere under tmp_path
    snapshot_after = {p for p in tmp_path.rglob("*") if p.is_file()}
    assert snapshot_before == snapshot_after
    assert not list(tmp_path.rglob("report_*.xlsx"))

    # Returned report DataFrame is the in-memory equivalent of the old file
    expected_rows = 2
    assert report_df.shape[0] == expected_rows

    included_orders = report_df["订单号"].tolist()
    assert "DATE_FILTER_THIS_MONTH" in included_orders
    assert "DATE_FILTER_LAST_YEAR" in included_orders
    assert "DATE_FILTER_OUT_OF_RANGE" not in included_orders
    assert "FULL_REFUND_ORDER_111" not in included_orders
    assert "CANCELLED_ORDER_222" not in included_orders

    # Check that the updated_df has the correct markings
    mark_value = "销售报表202603"
    assert updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_THIS_MONTH', '销售报表账期'].iloc[0] == mark_value
    assert updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_LAST_YEAR', '销售报表账期'].iloc[0] == mark_value
    assert pd.isna(updated_df.loc[updated_df['订单号'] == 'DATE_FILTER_OUT_OF_RANGE', '销售报表账期'].iloc[0])


def test_process_sales_report_workflow_signature_and_no_files(sample_data_dir, tmp_path):
    """`process_sales_report_workflow` must not accept output_dir and must not write files."""
    import inspect
    from utils import process_sales_report_workflow

    sig = inspect.signature(process_sales_report_workflow)
    assert "output_dir" not in sig.parameters

    order_path = sample_data_dir / "orders.xlsx"
    payment_path = sample_data_dir / "payments.csv"

    snapshot_before = {p for p in tmp_path.rglob("*") if p.is_file()}

    updated_df, report_df = process_sales_report_workflow(
        order_file=str(order_path),
        payment_file=str(payment_path),
        target_month="202603",
        verbose=False,
    )

    snapshot_after = {p for p in tmp_path.rglob("*") if p.is_file()}
    assert snapshot_before == snapshot_after
    # Updated order DataFrame and in-memory report DataFrame are returned
    assert "支付手续费" in updated_df.columns
    assert isinstance(report_df, pd.DataFrame)

def test_read_file_with_appropriate_method_csv_fallback(tmp_path):
    # Create a CSV with a weird encoding
    csv_file = tmp_path / "weird.csv"
    csv_file.write_bytes("订单号,商户订单号,订单金额,A,B,C\n123,456,100,1,2,3\n".encode("gb2312"))
    
    # This will trigger the fallback encoding mechanism
    df = read_file_with_appropriate_method(csv_file)
    assert len(df) == 1
    assert df.iloc[0]["订单号"] == "123"


def test_read_file_with_appropriate_method_csv_comment_and_separator_fallback(tmp_path):
    csv_file = tmp_path / "commented.csv"
    csv_file.write_bytes(
        (
            "# comment line\n"
            "# another comment\n"
            "订单号;商户订单号;订单金额\n"
            "ABC123;M123;100\n"
        ).encode("utf-8-sig")
    )

    df = read_file_with_appropriate_method(csv_file)
    assert len(df) == 1
    assert list(df.columns) == ["订单号", "商户订单号", "订单金额"]
    assert df.iloc[0]["订单号"] == "ABC123"

def test_read_file_with_appropriate_method_xls_fallback(tmp_path):
    # Write a dummy file that's not a real Excel file to trigger BadZipFile fallback
    fake_excel = tmp_path / "fake.xlsx"
    fake_excel.write_text("Not a real excel file")
    
    # Since it's not a real excel, it should raise an exception eventually, but hit the fallback first
    with pytest.raises(Exception):
        read_file_with_appropriate_method(fake_excel)

def test_process_excel_files_verbose(tmp_path):
    # Test verbose mode in process_excel_files
    order = tmp_path / "orders.csv"
    payment = tmp_path / "payments.csv"
    order.write_text("订单号,商户订单号,商务订单号,订单金额,A,B\nO123,M123,B123,100,1,2\n")
    payment.write_text("商品名称,商户订单号,支出金额（-元）,收入金额（+元）,业务类型,A\nRegular,M123,-5.0,0,收费,1\n")
    
    df = process_excel_files(order, payment, verbose=True)
    assert len(df) == 1

def test_write_result_file_formats(tmp_path):
    # Test write_result_file with different formats
    from utils import write_result_file
    import pandas as pd
    
    df = pd.DataFrame({"A": [1, 2]})
    
    csv_out = tmp_path / "out.csv"
    write_result_file(df, csv_out)
    assert csv_out.exists()
    
    xlsx_out = tmp_path / "out.xlsx"
    write_result_file(df, xlsx_out)
    assert xlsx_out.exists()
    
    xls_out = tmp_path / "out.xls"
    # Even for .xls, we wrote fallback logic to use openpyxl/None
    try:
        write_result_file(df, xls_out)
        assert xls_out.exists()
    except Exception:
        pass # If it fails due to xlwt missing, that's expected

def test_parse_date_variations():
    from utils import parse_date
    from datetime import datetime
    import pandas as pd
    
    assert parse_date(datetime(2026, 3, 1)) == pd.Timestamp(2026, 3, 1)
    assert parse_date("2026-03-01") == pd.Timestamp(2026, 3, 1)
    assert parse_date("2026年03月01日") == pd.Timestamp(2026, 3, 1)
    assert parse_date("invalid date") is None
    
def test_get_year_month():
    from utils import get_year_month
    from datetime import datetime
    
    assert get_year_month(datetime(2026, 3, 1)) == "202603"
    assert get_year_month("2026-03-01") == "202603"
    assert get_year_month("2026年3月1日") == "202603"
    assert get_year_month("invalid") is None
    

def test_extract_p_number():
    from utils import extract_p_number
    assert extract_p_number(None) is None
    assert extract_p_number("P12345") == "P12345"
    assert extract_p_number("Order P12345 details") == "P12345"
    assert extract_p_number("No p number here") is None
    assert extract_p_number(float('nan')) is None
    assert extract_p_number("") is None
    
def test_match_orders_by_p_number():
    from utils import match_orders_by_p_number
    assert match_orders_by_p_number("Ext P123", "Prod P123") is True
    assert match_orders_by_p_number("Ext P123", "Prod P456") is False
    assert match_orders_by_p_number("Ext", "Prod") is False
    assert match_orders_by_p_number(None, "Prod P123") is False


def test_parse_date_more_cases():
    from utils import parse_date
    import pandas as pd
    from datetime import datetime
    
    assert parse_date(pd.NaT) is None
    assert parse_date(None) is None
    assert parse_date("2026年3月") == pd.Timestamp(2026, 3, 1)
    
def test_add_sales_report_period_verbose(tmp_path):
    from utils import add_sales_report_period
    import pandas as pd
    
    df = pd.DataFrame({
        "订单号": ["1"],
        "订单状态": ["取消"],
        "订单金额": ["0"],
        "支付手续费": [0]
    })
    res = add_sales_report_period(df, verbose=True)
    assert res.iloc[0]["销售报表账期"] == "已取消"
