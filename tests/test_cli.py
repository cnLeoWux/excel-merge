"""
测试 CLI 功能

测试场景：
1. 基本匹配模式 (--match-only)
2. 基本标注模式 (--mark-only)
3. 默认模式（匹配+标注）
4. JSON 输出格式
5. 退出码验证
6. 错误处理（文件不存在、参数错误）
"""

import pytest
import subprocess
import sys
import json
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from tests.sample_data_generator import (
    data_generator,
    create_exact_match_scenario,
    create_refund_scenario,
    create_full_refund_scenario,
    create_cancelled_scenario,
)


class TestCLIBasicMatching:
    """测试 CLI 基本匹配模式"""

    def test_match_only_mode(self):
        """测试 --match-only 模式"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 0, f"CLI failed with: {result.stderr}"

            # 解析 JSON 输出
            output = json.loads(result.stdout)
            assert output["ok"] is True
            assert "statistics" in output["data"]
            assert output["data"]["statistics"]["matched_rows"] == 1
            assert output["data"]["statistics"]["total_rows"] == 1

    def test_match_only_updates_payment_fee(self):
        """测试 --match-only 填充支付手续费列"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # 读取结果文件，验证支付手续费已被填充
            import pandas as pd
            result_df = pd.read_excel(order_file)

            assert result_df.iloc[0]["支付手续费"] == -2.50


class TestCLIBasicMarking:
    """测试 CLI 基本标注模式"""

    def test_mark_only_mode(self):
        """测试 --mark-only 模式"""
        with data_generator() as gen:
            # 创建全退场景
            order_file = create_full_refund_scenario(gen, order_no="FULLREFUNDCLI001")

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(order_file), "202602",  # payment_file 用 order_file 代替
                    "--mark-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # --mark-only 模式下，支付文件实际上不会被使用（不进行匹配）
            # 所以即使支付文件不匹配也不会影响结果
            assert result.returncode == 0, f"CLI failed with: {result.stderr}"

    def test_mark_only_updates_sales_period(self):
        """测试 --mark-only 填充销售报表账期列"""
        with data_generator() as gen:
            order_file = create_full_refund_scenario(gen, order_no="FULLREFUNDCLI002")

            subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(order_file), "202602",
                    "--mark-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # 读取结果文件，验证销售报表账期已被填充
            import pandas as pd
            result_df = pd.read_excel(order_file)

            # 全退场景：两行都应该被标记
            for idx, row in result_df.iterrows():
                assert row["销售报表账期"] == "全退"


class TestCLIDefaultMode:
    """测试 CLI 默认模式（匹配+标注）"""

    def test_default_mode_both_matching_and_marking(self):
        """测试默认模式同时执行匹配和标注"""
        with data_generator() as gen:
            # 创建既有匹配又有标注的场景
            orders = [
                # 正单 - 可以匹配
                {"订单号": "40250702110303185340xx", "外部订单号": "", "订单金额": 100.0},
                # 全退场景 - 需要标注
                {"订单号": "FULLREFUND003", "外部订单号": "", "订单金额": 100.0},
                {"订单号": "FULLREFUND003", "外部订单号": "", "订单金额": -100.0},
            ]
            order_file = gen.create_order_file(orders)

            payments = [
                {"商户订单号": "40250702110303185340yy", "商品名称": "测试", "业务类型": "收费", "支出金额（-元）": -2.50},
            ]
            payment_file = gen.create_payment_file(payments)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 0, f"CLI failed with: {result.stderr}"

            output = json.loads(result.stdout)
            assert output["ok"] is True
            assert "matched_rows" in output["data"]["statistics"]
            assert "marked_rows" in output["data"]["statistics"]


class TestCLIJsonOutput:
    """测试 CLI JSON 输出格式"""

    def test_json_output_format(self):
        """测试 JSON 输出格式正确"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # 验证可以解析为 JSON
            output = json.loads(result.stdout)

            # 验证信封格式
            assert "ok" in output
            assert "data" in output
            assert "error" in output

            if output["ok"]:
                assert "statistics" in output["data"]
                stats = output["data"]["statistics"]
                assert "total_rows" in stats
                assert "matched_rows" in stats
                assert "match_rate" in stats
            else:
                assert output["error"] is not None
                assert "code" in output["error"]
                assert "message" in output["error"]


class TestCLIExitCodes:
    """测试 CLI 退出码"""

    def test_exit_code_success(self):
        """成功时退出码为 0"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 0

    def test_exit_code_file_not_found(self):
        """文件不存在时退出码为 3"""
        result = subprocess.run(
            [
                sys.executable, "cli.py",
                "nonexistent.xlsx", "payment.xlsx", "202602",
                "--json", "--quiet"
            ],
            capture_output=True,
            text=True,
            cwd=str(Path(__file__).parent.parent)
        )

        assert result.returncode == 3
        output = json.loads(result.stdout)
        assert output["ok"] is False
        assert output["error"]["code"] == "file_not_found"

    def test_exit_code_usage_error(self):
        """参数错误时退出码为 2"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            # 无效的 target_month 格式
            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "invalid",
                    "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 2
            output = json.loads(result.stdout)
            assert output["ok"] is False
            assert output["error"]["code"] == "usage_error"

    def test_exit_code_invalid_month(self):
        """无效月份时退出码为 2"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            # 月份超出范围
            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202613",  # 13月
                    "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 2

    def test_match_only_invalid_month_is_usage_error(self):
        """显式 reduced workflow 也拒绝无效 target_month"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202613",
                    "--match-only", "--json", "--quiet",
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent),
            )

            assert result.returncode == 2
            output = json.loads(result.stdout)
            assert output["ok"] is False
            assert output["error"]["code"] == "usage_error"


class TestCLIErrorHandling:
    """测试 CLI 错误处理"""

    def test_mutually_exclusive_flags(self):
        """测试 --match-only 和 --mark-only 互斥"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--mark-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # argparse 应该报错并以退出码 2 退出
            assert result.returncode == 2


class TestCLIQuietMode:
    """测试 CLI 静默模式"""

    def test_quiet_mode_no_stdout(self):
        """测试静默模式不输出到 stdout"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file), "202602",
                    "--match-only", "--json", "--quiet"
                ],
                capture_output=True,
                text=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # stdout 应该只包含 JSON
            try:
                json.loads(result.stdout)
            except json.JSONDecodeError:
                pytest.fail(f"stdout should be JSON only, got: {result.stdout}")


class TestCLIValidation:
    """测试 CLI 参数验证"""

    def test_target_month_format_validation(self):
        """测试 target_month 格式验证"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(gen)

            # 测试各种无效格式
            invalid_months = [
                "12345",     # 5位
                "1234567",   # 7位
                "abcdef",    # 非数字
                "202600",    # 月份为00
                "202613",    # 月份为13
                "201999",    # 月份为99
                "2019",      # 4位
            ]

            for invalid_month in invalid_months:
                result = subprocess.run(
                    [
                        sys.executable, "cli.py",
                        str(order_file), str(payment_file), invalid_month,
                        "--json", "--quiet"
                    ],
                    capture_output=True,
                    text=True,
                    cwd=str(Path(__file__).parent.parent)
                )

                assert result.returncode == 2, f"Expected exit code 2 for month {invalid_month}, got {result.returncode}"


class TestCLIInteractiveInput:
    """测试 CLI 交互式输入"""

    def test_missing_target_month_prompts_interactive_input(self):
        """不提供 target_month 时应提示用户输入"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            # 模拟用户输入 202602
            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file),
                    "--json", "--quiet"
                ],
                input="202602\n",
                text=True,
                capture_output=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # 应该成功执行完整工作流
            assert result.returncode == 0, f"Expected exit code 0, got {result.stderr}"
            output = json.loads(result.stdout)
            assert output["ok"] is True

    def test_interactive_input_with_valid_month(self):
        """交互输入有效的 target_month"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            # 用户输入 202603
            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file),
                    "--json"
                ],
                input="202603\n",
                text=True,
                capture_output=True,
                cwd=str(Path(__file__).parent.parent)
            )

            assert result.returncode == 0
            output = json.loads(result.stdout)
            assert output["ok"] is True

    def test_interactive_input_with_empty_then_valid(self):
        """用户先回车空输入，再输入有效值"""
        with data_generator() as gen:
            order_file, payment_file = create_exact_match_scenario(
                gen,
                order_no="40250702110303185340",
                amount=100.0,
                fee=-2.50
            )

            # 先空输入，再输入有效值
            result = subprocess.run(
                [
                    sys.executable, "cli.py",
                    str(order_file), str(payment_file),
                    "--json", "--quiet"
                ],
                input="\n202602\n",
                text=True,
                capture_output=True,
                cwd=str(Path(__file__).parent.parent)
            )

            # 应该继续等待直到获得有效输入
            assert result.returncode == 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
