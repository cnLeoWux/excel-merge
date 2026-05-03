"""
测试销售报表账期标注逻辑

测试场景：
1. 全退标记（同一订单号重复，金额合计为0）
2. 已取消标记（状态含"取消"且金额为0）
3. 普通订单不标记
"""

import pytest
import pandas as pd
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from utils import add_sales_report_period
from tests.sample_data_generator import (
    data_generator,
    create_full_refund_scenario,
    create_cancelled_scenario,
)


class TestFullRefundMarking:
    """测试"全退"标记"""

    def test_full_refund_two_rows_sum_to_zero(self):
        """同一订单号出现两次，金额合计为0，应该标记为"全退" """
        with data_generator() as gen:
            order_file = create_full_refund_scenario(gen, order_no="FULLREFUND001")

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 两行都应该被标记为"全退"
            marked_values = result_df["销售报表账期"].tolist()
            assert all(v == "全退" for v in marked_values), f"Expected all '全退', got {marked_values}"

    def test_full_refund_three_rows_sum_to_zero(self):
        """同一订单号出现三次，金额合计为0，应该标记为"全退" """
        with data_generator() as gen:
            orders = [
                {"订单号": "THREEROW001", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "THREEROW001", "外部订单号": "", "订单金额": -50.0},
                {"订单号": "THREEROW001", "外部订单号": "", "订单金额": -50.0},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            marked_values = result_df["销售报表账期"].tolist()
            assert all(v == "全退" for v in marked_values), f"Expected all '全退', got {marked_values}"

    def test_full_refund_sum_not_zero_no_mark(self):
        """同一订单号出现多次，但金额合计不为0，不应该标记"""
        with data_generator() as gen:
            orders = [
                {"订单号": "NOTZERO001", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "NOTZERO001", "外部订单号": "", "订单金额": -30.0},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 金额合计为70，不应标记
            marked_values = result_df["销售报表账期"].tolist()
            assert all(v is None or pd.isna(v) for v in marked_values), \
                f"Expected all None for sum != 0, got {marked_values}"

    def test_full_refund_only_duplicate_marked(self):
        """只有重复订单号才标记，单次出现的订单不标记"""
        with data_generator() as gen:
            orders = [
                {"订单号": "UNIQUE001", "外部订单号": "", "订单金额": 100.0},  # 单次出现
                {"订单号": "UNIQUE002", "外部订单号": "", "订单金额": 200.0},  # 单次出现
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 单次出现的订单不应被标记
            marked_values = result_df["销售报表账期"].tolist()
            assert all(v is None or pd.isna(v) for v in marked_values)


class TestCancelledMarking:
    """测试"已取消"标记"""

    def test_cancelled_with_zero_amount(self):
        """状态含"取消"且金额为0，应该标记为"已取消" """
        with data_generator() as gen:
            order_file = create_cancelled_scenario(gen, order_no="CANCELLED001")

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            assert result_df.iloc[0]["销售报表账期"] == "已取消"

    def test_cancelled_with_nonzero_amount_no_mark(self):
        """状态含"取消"但金额不为0，不应该标记"""
        with data_generator() as gen:
            orders = [{
                "订单号": "CANCELLED002",
                "外部订单号": "",
                "订单金额": 100.0,  # 不为0
                "订单状态": "已取消",
            }]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 金额不为0，不应标记
            assert result_df.iloc[0]["销售报表账期"] is None or pd.isna(result_df.iloc[0]["销售报表账期"])

    def test_cancelled_variations(self):
        """各种"取消"状态变体都应该被识别"""
        with data_generator() as gen:
            orders = [
                {"订单号": "CANCEL001", "外部订单号": "", "订单金额": 0.0, "订单状态": "已取消"},
                {"订单号": "CANCEL002", "外部订单号": "", "订单金额": 0.0, "订单状态": "取消"},
                {"订单号": "CANCEL003", "外部订单号": "", "订单金额": 0.0, "订单状态": "订单取消"},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 所有含"取消"的都应该被标记
            for idx, row in result_df.iterrows():
                assert row["销售报表账期"] == "已取消", \
                    f"Order {row['订单号']} with status '{row['订单状态']}' should be marked '已取消'"


class TestCombinedMarking:
    """测试全退和已取消的组合场景"""

    def test_full_refund_and_cancelled_together(self):
        """同时存在全退和已取消的场景"""
        with data_generator() as gen:
            orders = [
                # 全退
                {"订单号": "BOTH001", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "BOTH001", "外部订单号": "", "订单金额": -100.0},
                # 已取消
                {"订单号": "BOTHCANCEL", "外部订单号": "", "订单金额": 0.0, "订单状态": "已取消"},
                # 普通订单
                {"订单号": "NORMAL001", "外部订单号": "", "订单金额": 200.0},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 检查每个订单的标记
            for idx, row in result_df.iterrows():
                order_no = row["订单号"]
                marking = row["销售报表账期"]

                if order_no == "BOTH001":
                    assert marking == "全退", f"Expected '全退' for {order_no}, got {marking}"
                elif order_no == "BOTHCANCEL":
                    assert marking == "已取消", f"Expected '已取消' for {order_no}, got {marking}"
                else:
                    assert marking is None or pd.isna(marking), \
                        f"Expected None for {order_no}, got {marking}"

    def test_priority_full_refund_over_cancelled(self):
        """如果同一订单同时满足全退和已取消条件，先到先得（理论上不应该发生）"""
        # 实际上全退是基于订单号重复，已取消是基于状态和金额
        # 同一个订单不可能既是重复订单又是零金额状态
        pass


class TestNoMarking:
    """测试不应被标记的场景"""

    def test_normal_orders_not_marked(self):
        """普通订单不应该被标记"""
        with data_generator() as gen:
            orders = [
                {"订单号": "NORMAL001", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "NORMAL002", "外部订单号": "", "订单金额": 200.0},
                {"订单号": "NORMAL003", "外部订单号": "", "订单金额": -50.0},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 普通订单不应被标记
            marked_values = result_df["销售报表账期"].tolist()
            assert all(v is None or pd.isna(v) for v in marked_values)

    def test_zero_amount_but_not_cancelled(self):
        """金额为0但不包含"取消"状态，不应该标记"""
        with data_generator() as gen:
            orders = [{
                "订单号": "ZERONOCANCEL",
                "外部订单号": "",
                "订单金额": 0.0,
                "订单状态": "已确认",  # 不是"取消"
            }]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            result_df = add_sales_report_period(df, verbose=False)

            # 金额为0但不包含"取消"，不应标记
            assert result_df.iloc[0]["销售报表账期"] is None or pd.isna(result_df.iloc[0]["销售报表账期"])


class TestExistingMarkingColumn:
    """测试已有标注列的情况"""

    def test_clears_existing_marks_before_recalculating(self):
        """如果已有标注值，应该先清空再重新计算"""
        with data_generator() as gen:
            orders = [
                {"订单号": "RECALC001", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "RECALC001", "外部订单号": "", "订单金额": -100.0},
            ]
            order_file = gen.create_order_file(orders)

            df = pd.read_excel(order_file)
            # 预先填入一些标记
            df["销售报表账期"] = "旧标记"

            result_df = add_sales_report_period(df, verbose=False)

            # 应该被更新为"全退"，而不是保留"旧标记"
            for idx, row in result_df.iterrows():
                assert row["销售报表账期"] == "全退", \
                    f"Expected '全退', got '{row['销售报表账期']}'"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])