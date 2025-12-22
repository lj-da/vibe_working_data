"""
Session ID 管理器：用于记录和汇总每轮得到的 Session ID
"""

import os
import json
import datetime
import threading
from typing import List, Dict, Any


class SessionIDManager:
    """Session ID 管理器"""
    
    def __init__(self, result_dir: str, model_name: str, domain: str):
        self.result_dir = result_dir
        self.model_name = model_name
        self.domain = domain
        
        # 创建汇总文件路径
        self.summary_file = os.path.join(
            result_dir,
            "session_ids_summary.jsonl"
        )
        
        # 线程锁
        self.lock = threading.Lock()
        
        # 初始化汇总文件
        self._init_summary_file()
    
    def _init_summary_file(self):
        """初始化汇总文件"""
        if not os.path.exists(self.summary_file):
            with open(self.summary_file, "w", encoding="utf-8") as f:
                # 写入文件头信息
                header = {
                    "file_type": "session_ids_summary",
                    "model_name": self.model_name,
                    "domain": self.domain,
                    "created_at": datetime.datetime.now().isoformat(),
                    "description": "Session IDs generated during evaluation runs"
                }
                f.write(json.dumps(header, ensure_ascii=False) + "\n")
    
    def add_session_id(self, session_id: str, example_id: str, domain: str, 
                      result: float, stop_reason: str = "completed", 
                      steps: int = 0, additional_info: Dict = None):
        """
        添加 Session ID 到汇总文件
        
        Args:
            session_id: 会话ID
            example_id: 示例ID
            domain: 任务域
            result: 任务结果分数
            stop_reason: 停止原因
            steps: 执行步数
            additional_info: 额外信息
        """
        with self.lock:
            entry = {
                "timestamp": datetime.datetime.now().isoformat(),
                "session_id": session_id,
                "example_id": example_id,
                "domain": domain,
                "result": result,
                "stop_reason": stop_reason,
                "steps": steps,
                "model_name": self.model_name,
                "additional_info": additional_info or {}
            }
            
            with open(self.summary_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    
    def get_session_ids(self) -> List[Dict]:
        """获取所有 Session ID 记录"""
        session_ids = []
        
        if not os.path.exists(self.summary_file):
            return session_ids
        
        with open(self.summary_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('{"file_type":'):
                    try:
                        entry = json.loads(line)
                        session_ids.append(entry)
                    except json.JSONDecodeError:
                        continue
        
        return session_ids
    
    def get_recent_session_ids(self, count: int = 10) -> List[Dict]:
        """获取最近的 Session ID 记录"""
        all_sessions = self.get_session_ids()
        return all_sessions[-count:] if all_sessions else []
    
    def print_summary(self):
        """打印 Session ID 汇总信息"""
        sessions = self.get_session_ids()
        
        if not sessions:
            print("📝 暂无 Session ID 记录")
            return
        
        print(f"\n📋 Session ID 汇总 (共 {len(sessions)} 个)")
        print("=" * 80)
        
        # 按域分组统计
        domain_stats = {}
        for session in sessions:
            domain = session.get('domain', 'unknown')
            if domain not in domain_stats:
                domain_stats[domain] = {'count': 0, 'success': 0, 'total_score': 0}
            
            domain_stats[domain]['count'] += 1
            if session.get('result', 0) >= 1.0:
                domain_stats[domain]['success'] += 1
            domain_stats[domain]['total_score'] += session.get('result', 0)
        
        # 打印统计信息
        for domain, stats in domain_stats.items():
            success_rate = (stats['success'] / stats['count']) * 100 if stats['count'] > 0 else 0
            avg_score = stats['total_score'] / stats['count'] if stats['count'] > 0 else 0
            print(f"🏷️  {domain}: {stats['count']} 个任务, 成功率: {success_rate:.1f}%, 平均分数: {avg_score:.2f}")
        
        print("\n📝 最近的 Session ID:")
        recent_sessions = self.get_recent_session_ids(5)
        for session in recent_sessions:
            result = session.get('result', 0)
            status = "✅" if result >= 1.0 else "⚠️" if result >= 0.5 else "❌"
            print(f"  {status} {session['session_id']} - {session['example_id']} (分数: {result:.2f})")
        
        print(f"\n📁 完整记录文件: {self.summary_file}")
        print("=" * 80)
    
    def export_to_csv(self, output_file: str = None):
        """导出 Session ID 记录到 CSV 文件"""
        import csv
        
        if output_file is None:
            output_file = self.summary_file.replace('.jsonl', '.csv')
        
        sessions = self.get_session_ids()
        
        if not sessions:
            print("📝 暂无数据可导出")
            return
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['timestamp', 'session_id', 'example_id', 'domain', 'result', 'stop_reason', 'steps', 'model_name']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for session in sessions:
                writer.writerow({k: session.get(k, '') for k in fieldnames})
        
        print(f"📊 Session ID 记录已导出到: {output_file}")


def create_session_id_manager(args) -> SessionIDManager:
    """创建 Session ID 管理器"""
    return SessionIDManager(
        result_dir=args.result_dir,
        model_name=args.model,
        domain=args.domain
    )
