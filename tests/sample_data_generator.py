"""
测试数据生成器 - 为 Excel Merge Tool 生成测试数据

功能：
- 自动生成订单文件和支付文件
- 支持多种匹配场景：精确匹配、P-number匹配、连字符匹配、退单、零金额
- 自动清理生成的测试文件
"""

import os
import shutil
import tempfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple

import pandas as pd


class TestDataGenerator:
    """测试数据生成器"""

    def __init__(self, temp_dir: Optional[str] = None):
        """
        初始化测试数据生成器

        Args:
            temp_dir: 可选的临时目录路径。如果为 None，则使用系统临时目录。
        """
        if temp_dir:
            self.temp_dir = Path(temp_dir)
            self.temp_dir.mkdir(parents=True, exist_ok=True)
        else:
            self.temp_dir = Path(tempfile.mkdtemp(prefix="excel_merge_test_"))

        self.generated_files: List[Path] = []

    def _register_file(self, file_path: Path):
        """注册生成的文件，用于后续清理"""
        self.generated_files.append(file_path)
        return file_path

    def _make_unique_name(self, prefix: str, ext: str) -> str:
        """生成唯一的文件名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        return f"{prefix}_{timestamp}{ext}"

    def create_order_file(
        self,
        orders: List[Dict],
        filename: Optional[str] = None,
        as_csv: bool = False
    ) -> Path:
        """
        创建订单文件

        Args:
            orders: 订单数据列表，每项包含：
                    - 订单号: str
                    - 外部订单号: str (可选)
                    - 订单金额: float
                    - 订单状态: str (可选, 默认 "已确认")
                    - 出行日期: str (可选, 格式 YYYY-MM-DD)
            filename: 文件名（可选，自动生成）
            as_csv: 是否创建为 CSV 格式

        Returns:
            Path: 创建的文件路径
        """
        if filename is None:
            ext = ".csv" if as_csv else ".xlsx"
            filename = self._make_unique_name("order", ext)

        file_path = self.temp_dir / filename

        # 构建 DataFrame
        data = []
        for order in orders:
            row = {
                "订单号": order.get("订单号", ""),
                "外部订单号": order.get("外部订单号", ""),
                "订单金额": order.get("订单金额", 0.0),
                "订单状态": order.get("订单状态", "已确认"),
                "出行日期": order.get("出行日期", ""),
                "支付手续费": "",  # 待填充列
                "销售报表账期": "",  # 待填充列
            }
            data.append(row)

        df = pd.DataFrame(data)

        # 写入文件
        if as_csv:
            df.to_csv(file_path, index=False, encoding="utf-8-sig")
        else:
            df.to_excel(file_path, index=False, engine="openpyxl")

        return self._register_file(file_path)

    def create_payment_file(
        self,
        payments: List[Dict],
        filename: Optional[str] = None,
        as_csv: bool = False
    ) -> Path:
        """
        创建支付文件

        Args:
            payments: 支付数据列表，每项包含：
                    - 商户订单号: str
                    - 商品名称: str (可选)
                    - 业务类型: str ("收费", "退费", "退款", "服务费")
                    - 支出金额: float (负数表示支出，正单使用)
                    - 收入金额: float (正数表示收入，退单使用)
            filename: 文件名（可选，自动生成）
            as_csv: 是否创建为 CSV 格式

        Returns:
            Path: 创建的文件路径
        """
        if filename is None:
            ext = ".csv" if as_csv else ".xlsx"
            filename = self._make_unique_name("payment", ext)

        file_path = self.temp_dir / filename

        # 构建 DataFrame
        data = []
        for payment in payments:
            row = {
                "商户订单号": payment.get("商户订单号", ""),
                "商品名称": payment.get("商品名称", ""),
                "业务类型": payment.get("业务类型", "收费"),
                "支出金额（-元）": payment.get("支出金额（-元）", 0.0),
                "收入金额（+元）": payment.get("收入金额（+元）", 0.0),
            }
            data.append(row)

        df = pd.DataFrame(data)

        # 写入文件
        if as_csv:
            df.to_csv(file_path, index=False, encoding="utf-8-sig")
        else:
            df.to_excel(file_path, index=False, engine="openpyxl")

        return self._register_file(file_path)

    def cleanup(self):
        """清理所有生成的测试文件"""
        for file_path in self.generated_files:
            if file_path.exists():
                file_path.unlink()
        self.generated_files.clear()

    def __del__(self):
        """析构时确保清理"""
        self.cleanup()

    def get_temp_dir(self) -> Path:
        """获取临时目录路径"""
        return self.temp_dir


@contextmanager
def data_generator():
    """
    上下文管理器：创建测试数据生成器，使用后自动清理

    使用方式：
        with data_generator() as gen:
            order_file = gen.create_order_file([...])
            payment_file = gen.create_payment_file([...])
            # 执行测试...
        # 测试完成后自动清理
    """
    gen = TestDataGenerator()
    try:
        yield gen
    finally:
        gen.cleanup()


def create_exact_match_scenario(
    gen: TestDataGenerator,
    order_no: str = "40250702110303185340",
    amount: float = 100.0,
    fee: float = -2.50,
    as_csv: bool = False
) -> Tuple[Path, Path]:
    """
    创建精确匹配场景

    Args:
        gen: TestDataGenerator 实例
        order_no: 订单号（前20字符匹配）
        amount: 订单金额（正数=正单）
        fee: 支付手续费（负数表示支出）
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path]: (order_file, payment_file)
    """
    # 订单文件：20字符订单号
    orders = [{
        "订单号": order_no + "xx",  # 补齐到超过20字符
        "外部订单号": "",
        "订单金额": amount,
        "订单状态": "已确认",
    }]

    # 支付文件：相同的20字符前缀
    payments = [{
        "商户订单号": order_no + "yy",  # 相同20字符前缀
        "商品名称": "测试商品",
        "业务类型": "收费" if amount > 0 else "退费",
        "支出金额（-元）": fee if amount > 0 else 0.0,
        "收入金额（+元）": abs(fee) if amount < 0 else 0.0,
    }]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    return order_file, payment_file


def create_pnumber_match_scenario(
    gen: TestDataGenerator,
    p_number: str = "P2507021103060001",
    order_amount: float = 100.0,
    income: float = 1.50,
    as_csv: bool = False
) -> Tuple[Path, Path]:
    """
    创建 P-number 匹配场景

    Args:
        gen: TestDataGenerator 实例
        p_number: P-number（如 P2507021103060001）
        order_amount: 订单金额（正数=正单，负数=退单）
        income: 收入金额（退单时使用）
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path]: (order_file, payment_file)
    """
    # 订单文件：使用 P-number 作为外部订单号
    orders = [{
        "订单号": "SHORT001",  # 短订单号，不足20字符
        "外部订单号": p_number,
        "订单金额": order_amount,
        "订单状态": "已确认",
    }]

    # 支付文件：商品名称中包含 P-number
    payments = [{
        "商户订单号": "DIFFERENT001",
        "商品名称": f"测试商品-{p_number}",  # P-number 在连字符后
        "业务类型": "收费" if order_amount > 0 else "退费",
        "支出金额（-元）": -1.50 if order_amount > 0 else 0.0,
        "收入金额（+元）": income if order_amount < 0 else 0.0,
    }]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    return order_file, payment_file


def create_hyphen_match_scenario(
    gen: TestDataGenerator,
    hyphen_part: str = "H12345",
    order_amount: float = 100.0,
    fee: float = -1.00,
    as_csv: bool = False
) -> Tuple[Path, Path]:
    """
    创建连字符匹配场景

    Args:
        gen: TestDataGenerator 实例
        hyphen_part: 连字符后的部分
        order_amount: 订单金额
        fee: 手续费
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path]: (order_file, payment_file)
    """
    orders = [{
        "订单号": "ORDER" + hyphen_part,
        "外部订单号": hyphen_part,
        "订单金额": order_amount,
        "订单状态": "已确认",
    }]

    payments = [{
        "商户订单号": "DIFFERENT002",
        "商品名称": f"商品名称-{hyphen_part}",
        "业务类型": "收费",
        "支出金额（-元）": fee,
        "收入金额（+元）": 0.0,
    }]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    return order_file, payment_file


def create_refund_scenario(
    gen: TestDataGenerator,
    order_no: str = "40250702110303185340",
    refund_amount: float = -50.0,
    income: float = 1.20,
    as_csv: bool = False
) -> Tuple[Path, Path]:
    """
    创建退单场景

    Args:
        gen: TestDataGenerator 实例
        order_no: 订单号（应为基础20字符，确保前后缀不同但前20字符相同）
        refund_amount: 退款金额（负数）
        income: 收入金额（正值）
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path]: (order_file, payment_file)
    """
    orders = [{
        "订单号": order_no + "01",
        "外部订单号": "",
        "订单金额": refund_amount,
        "订单状态": "已退款",
    }]

    payments = [{
        "商户订单号": order_no + "02",
        "商品名称": "退款商品",
        "业务类型": "退费",
        "支出金额（-元）": 0.0,
        "收入金额（+元）": income,
    }]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    return order_file, payment_file


def create_zero_amount_scenario(
    gen: TestDataGenerator,
    order_no: str = "40250700999999999",
    as_csv: bool = False
) -> Tuple[Path, Path]:
    """
    创建零金额场景（应跳过匹配）

    Args:
        gen: TestDataGenerator 实例
        order_no: 订单号
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path]: (order_file, payment_file)
    """
    orders = [{
        "订单号": order_no + "xx",
        "外部订单号": "",
        "订单金额": 0.0,
        "订单状态": "已取消",
    }]

    payments = [{
        "商户订单号": order_no + "yy",
        "商品名称": "取消商品",
        "业务类型": "收费",
        "支出金额（-元）": -5.00,
        "收入金额（+元）": 0.0,
    }]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    return order_file, payment_file


def create_full_refund_scenario(
    gen: TestDataGenerator,
    order_no: str = "FULLREFUND001",
    as_csv: bool = False
) -> Path:
    """
    创建"全退"场景（同一订单号多次出现，金额合计为0）

    Args:
        gen: TestDataGenerator 实例
        order_no: 订单号
        as_csv: 是否创建 CSV 格式

    Returns:
        Path: 创建的订单文件路径
    """
    orders = [
        {
            "订单号": order_no,
            "外部订单号": "",
            "订单金额": 100.0,
            "订单状态": "已确认",
        },
        {
            "订单号": order_no,
            "外部订单号": "",
            "订单金额": -100.0,
            "订单状态": "已退款",
        },
    ]

    return gen.create_order_file(orders, as_csv=as_csv)


def create_cancelled_scenario(
    gen: TestDataGenerator,
    order_no: str = "CANCELLED001",
    as_csv: bool = False
) -> Path:
    """
    创建"已取消"场景（状态含"取消"且金额为0）

    Args:
        gen: TestDataGenerator 实例
        order_no: 订单号
        as_csv: 是否创建 CSV 格式

    Returns:
        Path: 创建的订单文件路径
    """
    orders = [{
        "订单号": order_no,
        "外部订单号": "",
        "订单金额": 0.0,
        "订单状态": "已取消",
    }]

    return gen.create_order_file(orders, as_csv=as_csv)


def create_mixed_scenario(
    gen: TestDataGenerator,
    as_csv: bool = False
) -> Tuple[Path, Path, Dict]:
    """
    创建混合场景（包含多种匹配类型的订单和支付文件）

    Args:
        gen: TestDataGenerator 实例
        as_csv: 是否创建 CSV 格式

    Returns:
        Tuple[Path, Path, Dict]: (order_file, payment_file, expected_results)
            expected_results 包含每个订单的预期匹配结果
    """
    orders = [
        # 正单 - 精确匹配
        {
            "订单号": "40250702110303185340xx",
            "外部订单号": "",
            "订单金额": 100.0,
            "订单状态": "已确认",
        },
        # 退单 - P-number匹配
        {
            "订单号": "SHORT001",
            "外部订单号": "P2507021103060001",
            "订单金额": -50.0,
            "订单状态": "已退款",
        },
        # 正单 - 连字符匹配
        {
            "订单号": "ORDERH12345",
            "外部订单号": "H12345",
            "订单金额": 200.0,
            "订单状态": "已确认",
        },
        # 零金额订单 - 应跳过
        {
            "订单号": "ZEROTEST001",
            "外部订单号": "",
            "订单金额": 0.0,
            "订单状态": "已取消",
        },
        # 无匹配订单
        {
            "订单号": "NOMATCH001",
            "外部订单号": "",
            "订单金额": 150.0,
            "订单状态": "已确认",
        },
    ]

    payments = [
        # 精确匹配：订单号前20字符
        {
            "商户订单号": "40250702110303185340yy",
            "商品名称": "测试商品A",
            "业务类型": "收费",
            "支出金额（-元）": -2.50,
            "收入金额（+元）": 0.0,
        },
        # P-number匹配：商品名称含 P2507021103060001
        {
            "商户订单号": "DIFFERENT001",
            "商品名称": "商品-P2507021103060001",
            "业务类型": "退费",
            "支出金额（-元）": 0.0,
            "收入金额（+元）": 1.20,
        },
        # 连字符匹配：商品名称含 H12345
        {
            "商户订单号": "DIFFERENT002",
            "商品名称": "商品名称-H12345",
            "业务类型": "收费",
            "支出金额（-元）": -3.00,
            "收入金额（+元）": 0.0,
        },
    ]

    order_file = gen.create_order_file(orders, as_csv=as_csv)
    payment_file = gen.create_payment_file(payments, as_csv=as_csv)

    # 预期结果
    expected = {
        "40250702110303185340xx": {"matched": True, "fee": -2.50},
        "SHORT001": {"matched": True, "fee": 1.20},
        "ORDERH12345": {"matched": True, "fee": -3.00},
        "ZEROTEST001": {"matched": True, "fee": 0.0},  # 零金额跳过
        "NOMATCH001": {"matched": False, "fee": None},
    }

    return order_file, payment_file, expected
