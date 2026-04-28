import pytest
from pathlib import Path
import pandas as pd
import shutil

@pytest.fixture(scope="session")
def sample_data_dir(tmp_path_factory):
    """
    Creates a temporary directory with sample data files for testing.
    Returns the Path object of the directory.
    """
    temp_dir = tmp_path_factory.mktemp("data")
    
    # --- Create Order File ---
    order_data = {
        "订单号": [
            "EXACT_MATCH_ORDER_12345",
            "EXACT_MATCH_REFUND_67890",
            "ORDER_WITH_PNUM_ABCDE",
            "ORDER_WITH_HYPHEN_FGHIJ",
            "FULL_REFUND_ORDER_111",
            "FULL_REFUND_ORDER_111",
            "CANCELLED_ORDER_222",
            "UNMATCHED_ORDER_333",
            "ZERO_AMOUNT_ORDER_444",
            "DATE_FILTER_THIS_MONTH",
            "DATE_FILTER_LAST_YEAR",
            "DATE_FILTER_OUT_OF_RANGE",
        ],
        "外部订单号": [
            "",
            "",
            "Some data with P12345",
            "some-external-id",
            "",
            "",
            "",
            "no-match-p-num",
            "",
            "",
            "",
            "",
        ],
        "订单金额": [
            100.0, -50.0, 200.0, 300.0, 88.0, -88.0, 0.0, 500.0, 0.0,
            150.0, 160.0, 170.0,
        ],
        "订单状态": [
            "已完成", "已完成", "已完成", "已完成", "已完成",
            "已退款", "已取消", "已完成", "已完成", "已完成",
            "已完成", "已完成",
        ],
        "出行日期": [
            "2020-01-15", "2020-01-16", "2020-01-17", "2020-01-18",
            "2020-01-19", "2020-01-20", "2020-01-21", "2020-01-22",
            "2020-01-23", "2026-03-10", "2025-04-10", "2024-01-01",
        ],
    }
    order_df = pd.DataFrame(order_data)
    order_df.to_excel(temp_dir / "orders.xlsx", index=False)
    order_df.to_csv(temp_dir / "orders.csv", index=False, encoding="utf-8-sig")

    # --- Create Payment File ---
    payment_data = {
        "商户订单号": [
            "EXACT_MATCH_ORDER_12345",
            "EXACT_MATCH_REFUND_67890",
            "some_other_id_1",
            "some_other_id_2",
            "some_other_id_3",
        ],
        "商品名称": [
            "some product",
            "some other product",
            "Product with P12345 info",
            "product-name-ends-with-some-external-id",
            "product for another order",
        ],
        "业务类型": ["收费", "退费", "服务费", "收费", "收费"],
        "收入金额（+元）": [0.0, 5.0, 0.0, 0.0, 0.0],
        "支出金额（-元）": [-10.0, 0.0, -20.0, -30.0, -99.9],
    }
    payment_df = pd.DataFrame(payment_data)
    payment_df.to_excel(temp_dir / "payments.xlsx", index=False)
    payment_df.to_csv(temp_dir / "payments.csv", index=False, encoding="utf-8-sig")

    return temp_dir
