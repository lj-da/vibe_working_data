#!/usr/bin/env python3
"""
Session ID 查看工具
用于查看和管理 Session ID 汇总
"""

import argparse
import os
import sys
from session_id_manager import SessionIDManager


def main():
    parser = argparse.ArgumentParser(description="Session ID 查看工具")
    parser.add_argument("--result_dir", type=str, default="./results", help="结果目录")
    parser.add_argument("--model", type=str, required=True, help="模型名称")
    parser.add_argument("--domain", type=str, default="all", help="任务域")
    parser.add_argument("--export_csv", action="store_true", help="导出为 CSV 文件")
    parser.add_argument("--recent", type=int, default=10, help="显示最近的 N 个 Session ID")
    
    args = parser.parse_args()
    
    try:
        # 创建 Session ID 管理器
        manager = SessionIDManager(
            result_dir=args.result_dir,
            model_name=args.model,
            domain=args.domain
        )
        
        # 显示汇总信息
        manager.print_summary()
        
        # 导出 CSV（如果请求）
        if args.export_csv:
            manager.export_to_csv()
        
        # 显示最近的 Session ID
        if args.recent > 0:
            print(f"\n📝 最近的 {args.recent} 个 Session ID:")
            recent_sessions = manager.get_recent_session_ids(args.recent)
            for i, session in enumerate(recent_sessions, 1):
                result = session.get('result', 0)
                status = "✅" if result >= 1.0 else "⚠️" if result >= 0.5 else "❌"
                print(f"  {i:2d}. {status} {session['session_id']} - {session['example_id']} (分数: {result:.2f})")
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
