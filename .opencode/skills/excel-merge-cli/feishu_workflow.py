#!/usr/bin/env python3
"""
Excel Merge - Feishu Workflow Helper

This script handles the complete Feishu workflow:
1. Download uploaded files from Feishu
2. Identify order vs payment files
3. Run excel-merge-cli
4. Upload processed file back to Feishu
5. Send result message to chat

Usage:
    python feishu_workflow.py --order-file /path/to/order.xlsx --payment-file /path/to/payment.xlsx --chat-id oc_xxx [--month 202602]
"""

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def run_cli(order_file: str, payment_file: str, month: str = None) -> dict:
    """Run excel-merge-cli and return parsed JSON output."""
    cmd = [
        "python", "cli.py",
        order_file,
        payment_file,
        "--json", "--quiet"
    ]
    
    if month:
        cmd.extend(["--month", month])
    
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent.parent)  # Project root
    )
    
    try:
        output = json.loads(result.stdout)
        return output
    except json.JSONDecodeError:
        return {
            "ok": False,
            "data": None,
            "error": {
                "code": "cli_error",
                "message": f"CLI failed with exit code {result.returncode}: {result.stderr}"
            }
        }


def identify_file_type(file_path: str) -> str:
    """Identify if file is order or payment by column names."""
    import pandas as pd
    
    try:
        df = pd.read_excel(file_path) if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
        columns = [str(c).strip() for c in df.columns]
        
        # Order file indicators
        if '订单号' in columns and '订单金额' in columns:
            return 'order'
        
        # Payment file indicators
        if any('商户' in c and '订单' in c for c in columns) or '商户订单号' in columns:
            return 'payment'
        
        # Check for amount columns
        if any('支出金额' in c or '收入金额' in c for c in columns):
            return 'payment'
        
        return 'unknown'
    except Exception as e:
        return f'error: {e}'


def main():
    parser = argparse.ArgumentParser(description='Excel Merge Feishu Workflow')
    parser.add_argument('--order-file', required=True, help='Path to order file')
    parser.add_argument('--payment-file', required=True, help='Path to payment file')
    parser.add_argument('--chat-id', required=True, help='Feishu chat ID to send result')
    parser.add_argument('--month', help='Target month for sales report (YYYYMM)')
    parser.add_argument('--auto-detect', action='store_true', help='Auto-detect file types and swap if needed')
    
    args = parser.parse_args()
    
    # Auto-detect file types if requested
    if args.auto_detect:
        order_type = identify_file_type(args.order_file)
        payment_type = identify_file_type(args.payment_file)
        
        # Swap if they seem reversed
        if order_type == 'payment' and payment_type == 'order':
            print("Detected reversed files, swapping...")
            args.order_file, args.payment_file = args.payment_file, args.order_file
        elif order_type == 'unknown' or payment_type == 'unknown':
            print(f"Warning: Could not identify file types (order={order_type}, payment={payment_type})")
    
    # Create processed copy
    order_path = Path(args.order_file)
    processed_path = order_path.parent / f"{order_path.stem}_processed{order_path.suffix}"
    shutil.copy(args.order_file, processed_path)
    
    print(f"Created processed copy: {processed_path}")
    
    # Run CLI
    result = run_cli(str(processed_path), args.payment_file, args.month)
    
    if result["ok"]:
        stats = result["data"]["statistics"]
        
        # Build success message
        message = f"""✅ Excel 合并完成！

📊 匹配统计：
   • 总订单数：{stats['total_rows']}
   • 成功匹配：{stats['matched_rows']}
   • 匹配率：{stats['match_rate']}"""
        
        if args.month:
            message += f"""

📝 销售报表：
   • 目标月份：{args.month}
   • 已标注账期信息"""
        
        message += "\n\n📎 已处理文件已保存，请手动上传到飞书"
        
        print(message)
        print(f"\nProcessed file: {processed_path}")
        
        # Return success info as JSON for automation
        output = {
            "success": True,
            "message": message,
            "processed_file": str(processed_path),
            "statistics": stats,
            "chat_id": args.chat_id
        }
        print(json.dumps(output, ensure_ascii=False))
        return 0
        
    else:
        error = result["error"]
        message = f"""❌ 处理失败：{error['message']}

请检查：
   • 文件格式是否正确 (.xlsx/.xls/.csv)
   • 订单文件是否包含"订单号"列
   • 支付文件是否包含"商户订单号"列
   • 文件是否被其他程序占用"""
        
        print(message)
        
        output = {
            "success": False,
            "message": message,
            "error": error,
            "chat_id": args.chat_id
        }
        print(json.dumps(output, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
