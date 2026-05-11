"""
CLI 主入口 - Excel Merge Tool
支持3种工作模式：
1. 有 target_month: 匹配 + 标注 + 日期筛选（完整工作流）
2. 无 target_month: 交互式输入或仅匹配
3. --match-only / --mark-only: 单独执行匹配或标注

处理前自动备份订单文件
"""

import json
import logging
import sys
from pathlib import Path

from utils import auto_backup
from workflow_service import (
    WorkflowError,
    run_mark_only,
    run_match_only,
    run_sales_report,
    validate_target_month_value,
)

# Exit codes
EXIT_SUCCESS = 0
EXIT_GENERAL_ERROR = 1
EXIT_USAGE_ERROR = 2
EXIT_FILE_NOT_FOUND = 3
EXIT_PROCESSING_ERROR = 4


def output_result(data=None, error=None, json_mode=False):
    """
    输出结果到 stdout

    JSON mode: 输出标准信封格式 {ok, data, error}
    Text mode: 输出人类可读消息
    """
    if json_mode:
        if error:
            result = {
                "ok": False,
                "data": None,
                "error": {
                    "code": error.get("code", "unknown_error"),
                    "message": error.get("message", "Unknown error"),
                },
            }
        else:
            result = {
                "ok": True,
                "data": data,
                "error": None,
            }
        print(json.dumps(result, ensure_ascii=False))
    else:
        # Text mode
        if error:
            print(f"错误: {error['message']}")

        if data:
            if "output_file" in data:
                print(f"订单文件已就地更新: {data['output_file']}")

            if "message" in data:
                print(data["message"])


def validate_target_month(target_month: str) -> bool:
    """
    验证 target_month 格式是否有效

    验证规则：
    1. 必须正好6位数字
    2. 年份范围：2020-2099
    3. 月份范围：01-12

    Args:
        target_month: 待验证的月份字符串

    Returns:
        bool: 验证是否通过
    """
    return validate_target_month_value(target_month)


def main_cli():
    """
    CLI 主入口函数

    工作流分叉：
    1. 有 target_month（无 --match-only 且无 --mark-only）: 完整工作流（匹配+标注+日期筛选）
    2. 无 target_month（无 --match-only 且无 --mark-only）: 交互式输入 target_month 或仅匹配
    3. --match-only: 仅执行匹配（填充支付手续费）
    4. --mark-only: 仅执行标注（标记销售报表账期）

    注意：--match-only 和 --mark-only 需要 target_month 为必填
    """
    import argparse

    parser = argparse.ArgumentParser(
        description="Excel Merge Tool - 订单与支付流水匹配工具"
    )

    # 位置参数
    parser.add_argument(
        "order_file", type=str, help="订单文件路径 (.xlsx, .xls, .csv)"
    )
    parser.add_argument(
        "payment_file", type=str, help="支付文件路径 (.xlsx, .xls, .csv)"
    )
    parser.add_argument(
        "target_month", type=str, nargs="?", default=None,
        help="目标月份 (格式: YYYYMM, 如 202602)。有值时执行完整工作流（含日期筛选）"
    )

    # 操作模式（互斥组）
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--match-only",
        action="store_true",
        help="仅执行订单匹配（填充支付手续费）",
    )
    mode_group.add_argument(
        "--mark-only",
        action="store_true",
        help="仅执行对账标注（标记销售报表账期）",
    )

    # 输出控制
    parser.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 格式输出结果到 stdout",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="详细日志模式 (-v=INFO, -vv=DEBUG)",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="静默模式（仅输出警告和错误）",
    )

    args = parser.parse_args()

    # ==================== 交互式输入 target_month ====================
    # 如果没有提供 target_month 且没有指定 --match-only 或 --mark-only
    # 则进入交互模式提示用户输入
    if args.target_month is None and not args.match_only and not args.mark_only:
        import sys as _sys
        if not args.json:
            print("提示：未指定目标月份，将执行完整工作流（匹配+标注+日期筛选）", file=_sys.stderr)
        while True:
            try:
                print("请输入目标月份 (YYYYMM格式，如 202602)，直接回车退出: ", end="", file=_sys.stderr, flush=True)
                user_input = input().strip()
                if user_input == "":
                    # 用户直接回车，询问是否退出
                    if not args.json:
                        print("确定退出吗？ (y/n): ", end="", file=_sys.stderr, flush=True)
                        confirm = input().strip().lower()
                        if confirm in ("y", "yes", ""):
                            output_result(
                                data={"message": "已取消"},
                                json_mode=args.json,
                            )
                            sys.exit(EXIT_SUCCESS)
                        else:
                            continue
                if user_input == "":
                    continue
                if validate_target_month(user_input):
                    args.target_month = user_input
                    if not args.json:
                        print(f"已选择目标月份: {args.target_month}", file=_sys.stderr)
                    break
                else:
                    error_msg = "target_month 格式无效，请输入6位数字 (YYYYMM格式)"
                    if not args.json:
                        print(error_msg, file=_sys.stderr)
                    else:
                        output_result(
                            error={"code": "usage_error", "message": error_msg},
                            json_mode=args.json,
                        )
                        sys.exit(EXIT_USAGE_ERROR)
            except EOFError:
                # stdin 关闭，直接退出
                output_result(
                    data={"message": "已取消（无输入）"},
                    json_mode=args.json,
                )
                sys.exit(EXIT_SUCCESS)
            except KeyboardInterrupt:
                print(file=_sys.stderr)
                output_result(
                    data={"message": "已取消"},
                    json_mode=args.json,
                )
                sys.exit(EXIT_SUCCESS)

    # ==================== 模式验证 ====================
    # --match-only 和 --mark-only 需要 target_month
    if (args.match_only or args.mark_only) and args.target_month is None:
        output_result(
            error={"code": "usage_error", "message": "--match-only 和 --mark-only 需要提供 target_month"},
            json_mode=args.json,
        )
        sys.exit(EXIT_USAGE_ERROR)

    if (args.match_only or args.mark_only) and args.target_month is not None and not validate_target_month(args.target_month):
        output_result(
            error={"code": "usage_error", "message": "target_month 格式无效"},
            json_mode=args.json,
        )
        sys.exit(EXIT_USAGE_ERROR)

    # ==================== 配置日志 ====================
    if args.quiet:
        log_level = logging.WARNING
    elif args.verbose >= 2:
        log_level = logging.DEBUG
    elif args.verbose >= 1:
        log_level = logging.INFO
    else:
        log_level = logging.INFO

    logging.basicConfig(
        level=log_level,
        format="%(message)s",
        stream=sys.stderr,
    )

    # 显示文件信息（非 JSON 非静默模式）
    if not args.json and not args.quiet:
        print(f"处理文件:")
        print(f"  订单文件: {args.order_file}")
        print(f"  支付文件: {args.payment_file}")
        if args.target_month:
            print(f"  目标月份: {args.target_month}")

    # ==================== 自动备份 ====================
    # CLI 保持当前备份契约：适配器在调用 service 前创建备份；service 只负责写回协调。
    backup_path = auto_backup(args.order_file)
    if not args.json and not args.quiet:
        print(f"  备份已创建: {Path(backup_path).name}")

    # ==================== 执行处理 ====================
    try:
        verbose = args.verbose >= 1 and not args.quiet

        if args.match_only:
            # 分支1: 仅匹配
            if not args.json and not args.quiet:
                print(f"\n执行仅匹配模式...")

            result = run_match_only(args.order_file, args.payment_file, verbose=verbose)
            stats = result.statistics or {}

            if not args.json and not args.quiet:
                print(
                    f"匹配完成: {stats.get('matched_rows', 0)}/{stats.get('total_rows', 0)} "
                    f"({stats.get('match_rate', '0.00%')})"
                )

            output_result(
                data={"output_file": result.output_file, "statistics": stats},
                json_mode=args.json,
            )

        elif args.mark_only:
            # 分支2: 仅标注
            if not args.json and not args.quiet:
                print(f"\n执行仅标注模式...")

            result = run_mark_only(args.order_file, verbose=verbose)
            stats = result.statistics or {}

            if not args.json and not args.quiet:
                print(f"标注完成: {stats.get('marked_rows', 0)}/{stats.get('total_rows', 0)} 行已标注")

            output_result(
                data={"output_file": result.output_file, "statistics": stats},
                json_mode=args.json,
            )

        elif args.target_month is not None:
            # 分支3: 完整工作流（有 target_month）
            # 执行匹配 + 标注 + 日期筛选
            if not args.json and not args.quiet:
                print(f"\n执行完整工作流（匹配+标注+日期筛选）...")

            result = run_sales_report(
                args.order_file, args.payment_file, args.target_month, verbose=verbose
            )
            stats = result.statistics or {}

            if not args.json and not args.quiet:
                print(
                    f"完成: 匹配 {stats.get('matched_rows', 0)}/{stats.get('total_rows', 0)} "
                    f"({stats.get('match_rate', '0.00%')}), 标注 {stats.get('marked_rows', 0)} 行"
                )

            output_result(
                data={"output_file": result.output_file, "statistics": stats},
                json_mode=args.json,
            )

        else:
            # 分支4: 仅匹配（无 target_month，无 flag）
            if not args.json and not args.quiet:
                print(f"\n执行仅匹配模式（无目标月份）...")

            result = run_match_only(args.order_file, args.payment_file, verbose=verbose)
            stats = result.statistics or {}

            if not args.json and not args.quiet:
                print(
                    f"匹配完成: {stats.get('matched_rows', 0)}/{stats.get('total_rows', 0)} "
                    f"({stats.get('match_rate', '0.00%')})"
                )

            output_result(
                data={"output_file": result.output_file, "statistics": stats},
                json_mode=args.json,
            )

        sys.exit(EXIT_SUCCESS)

    except WorkflowError as e:
        output_result(
            error={"code": e.code, "message": e.message},
            json_mode=args.json,
        )
        sys.exit(e.exit_code or EXIT_PROCESSING_ERROR)
    except Exception as e:
        output_result(
            error={"code": "processing_error", "message": str(e)},
            json_mode=args.json,
        )
        if not args.json:
            import traceback
            traceback.print_exc()
        sys.exit(EXIT_PROCESSING_ERROR)


if __name__ == "__main__":
    main_cli()
