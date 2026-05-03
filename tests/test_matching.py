"""
测试订单匹配逻辑

测试场景：
1. 精确匹配（20字符）
2. P-number 匹配
3. 连字符匹配
4. 业务类型校验
5. 零金额跳过
"""

import pytest
import pandas as pd
import sys
from pathlib import Path

# 将项目根目录添加到 sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils import (
    extract_p_number,
    match_orders_by_p_number,
    process_excel_files,
    read_file_with_appropriate_method,
)
from tests.sample_data_generator import (
    data_generator,
    create_exact_match_scenario,
    create_pnumber_match_scenario,
    create_hyphen_match_scenario,
    create_refund_scenario,
    create_zero_amount_scenario,
    create_mixed_scenario,
)


class TestExtractPNumber:
    """测试 P-number 提取函数"""

    def test_extract_p_number_with_valid_p_number(self):
        """提取有效的 P-number"""
        result = extract_p_number("P2507021103060001")
        assert result == "P2507021103060001"

    def test_extract_p_number_with_prefix(self):
        """提取带前缀的 P-number"""
        result = extract_p_number("ORDER-P2507021103060001")
        assert result == "P2507021103060001"

    def test_extract_p_number_no_match(self):
        """无 P-number 时返回 None"""
        assert extract_p_number("NO_P_NUMBER") is None
        assert extract_p_number("p12345") is None  # 小写不匹配

    def test_extract_p_number_with_none(self):
        """None 或 NaN 值返回 None"""
        import pandas as pd
        assert extract_p_number(None) is None
        assert extract_p_number(pd.NA) is None
        assert extract_p_number("") is None


class TestMatchOrdersByPNumber:
    """测试 P-number 匹配函数"""

    def test_match_with_same_p_number(self):
        """相同的 P-number 应该匹配"""
        assert match_orders_by_p_number("P2507021103060001", "商品-P2507021103060001") is True

    def test_match_with_different_p_number(self):
        """不同的 P-number 不应该匹配"""
        assert match_orders_by_p_number("P2507021103060001", "商品-P2507012326430003") is False

    def test_match_with_missing_external(self):
        """缺少外部订单号时不匹配"""
        assert match_orders_by_p_number(None, "商品-P2507021103060001") is False

    def test_match_with_missing_product(self):
        """缺少商品名称时不匹配"""
        assert match_orders_by_p_number("P2507021103060001", None) is False


class TestExactMatching:
    """测试精确匹配（20字符）"""

    def test_exact_match_same_first_20_chars(self):
        """订单号前20字符相同时应该匹配"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            # 验证匹配结果
            assert len(result_df) == 1
            assert result_df.iloc[0]["支付手续费"] == -2.50

    def test_exact_match_different_first_20_chars(self):
        """订单号前20字符不同时不应该匹配"""
        with data_generator() as gen:
            # 创建订单：订单号前20字符是 "40250702110303185340"
            orders = [{"订单号": "40250702110303185340xx", "外部订单号": "", "订单金额": 100.0}]
            order_file = gen.create_order_file(orders)

            # 创建支付：商户订单号前20字符是 "40250702110303185341"（最后一位不同）
            payments = [{"商户订单号": "40250702110303185341yy", "商品名称": "测试", "业务类型": "收费", "支出金额（-元）": -2.50}]
            payment_file = gen.create_payment_file(payments)

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            # 验证未匹配
            assert len(result_df) == 1
            assert result_df.iloc[0]["支付手续费"] is None or pd.isna(result_df.iloc[0]["支付手续费"])

    def test_exact_match_business_type_validation(self):
        """业务类型校验：正单应该匹配"收费"类型"""
        with data_generator() as gen:
            orders = [{"订单号": "40250702110303185340xx", "外部订单号": "", "订单金额": 100.0}]
            order_file = gen.create_order_file(orders)

            # 业务类型为"退费"（应该是退单匹配）
            payments = [{"商户订单号": "40250702110303185340yy", "商品名称": "测试", "业务类型": "退费", "支出金额（-元）": -2.50}]
            payment_file = gen.create_payment_file(payments)

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            # 正单不应该匹配退费类型
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is None or pd.isna(fee)


class TestPNumberMatching:
    """测试 P-number 匹配"""

    def test_pnumber_match_success(self):
        """P-number 匹配成功"""
        with data_generator() as gen:
            order_file, payment_file = create_pnumber_match_scenario(
                gen,
                p_number="P2507021103060001",
                order_amount=100.0,
                income=1.50
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            # P-number 匹配应该成功
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is not None and not pd.isna(fee)

    def test_pnumber_match_case_sensitive(self):
        """P-number 匹配区分大小写"""
        p_num = "P2507021103060001"
        assert extract_p_number(p_num) == p_num

        # 小写 p 不匹配
        assert extract_p_number("p2507021103060001") is None


class TestHyphenMatching:
    """测试连字符匹配"""

    def test_hyphen_match_success(self):
        """连字符匹配成功"""
        with data_generator() as gen:
            order_file, payment_file = create_hyphen_match_scenario(
                gen,
                hyphen_part="H12345",
                order_amount=100.0,
                fee=-1.00
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is not None and not pd.isna(fee)
            assert fee == -1.00


class TestRefundMatching:
    """测试退单匹配"""

    def test_refund_match_success(self):
        """退单匹配成功"""
        with data_generator() as gen:
            order_file, payment_file = create_refund_scenario(
                gen,
                order_no="40250702110303185340",
                refund_amount=-50.0,
                income=1.20
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is not None and not pd.isna(fee)
            assert fee == 1.20  # 退单：收入金额为正

    def test_refund_match_wrong_business_type(self):
        """退单不应匹配"收费"类型"""
        with data_generator() as gen:
            orders = [{"订单号": "4025070211030318534001", "外部订单号": "", "订单金额": -50.0}]
            order_file = gen.create_order_file(orders)

            payments = [{"商户订单号": "40250701123456789yy", "商品名称": "测试", "业务类型": "收费", "支出金额（-元）": -2.50}]
            payment_file = gen.create_payment_file(payments)

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            # 退单不应匹配收费类型
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is None or pd.isna(fee)


class TestZeroAmountSkipping:
    """测试零金额订单跳过"""

    def test_zero_amount_order_skipped(self):
        """零金额订单应跳过匹配，直接设 fee 为 0"""
        with data_generator() as gen:
            order_file, payment_file = create_zero_amount_scenario(
                gen,
                order_no="40250700999999999"
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            fee = result_df.iloc[0]["支付手续费"]
            assert fee == 0.0  # 零金额订单直接设为 0


class TestMixedMatching:
    """测试混合匹配场景"""

    def test_mixed_scenario(self):
        """混合场景：包含多种匹配类型"""
        with data_generator() as gen:
            order_file, payment_file, expected = create_mixed_scenario(gen)

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 5

            # 检查每个订单的匹配结果
            for idx, row in result_df.iterrows():
                order_no = row["订单号"]
                expected_result = expected.get(order_no, {})

                if expected_result.get("matched"):
                    expected_fee = expected_result.get("fee")
                    if expected_fee == 0.0:
                        assert row["支付手续费"] == 0.0
                    else:
                        assert row["支付手续费"] is not None and not pd.isna(row["支付手续费"])
                else:
                    # 无匹配
                    fee = row["支付手续费"]
                    assert fee is None or pd.isna(fee)


class TestCSVSupport:
    """测试 CSV 文件支持"""

    def test_csv_exact_matching(self):
        """CSV 文件：精确匹配"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50,
                as_csv=True
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            assert result_df.iloc[0]["支付手续费"] == -2.50

    def test_csv_pnumber_matching(self):
        """CSV 文件：P-number 匹配"""
        with data_generator() as gen:
            order_file, payment_file = create_pnumber_match_scenario(
                gen,
                p_number="P2507021103060001",
                order_amount=-50.0,
                income=1.20,
                as_csv=True
            )

            result_df = process_excel_files(str(order_file), str(payment_file), verbose=False)

            assert len(result_df) == 1
            fee = result_df.iloc[0]["支付手续费"]
            assert fee is not None and not pd.isna(fee)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])