#!/usr/bin/env python3
"""
简化版本的OSWorld成功率计算器
快速计算指定路径下的总体成功率
"""

import os
import glob
import argparse


def quick_calculate(results_path=None, model_name=None):
    """快速计算成功率"""
    if results_path is None:
        results_path = "/home/agent/code/OSWorld/results"
    
    if model_name is None:
        print("错误: 必须指定模型名称")
        return
    
    if not os.path.exists(results_path):
        print(f"错误: 路径不存在: {results_path}")
        return
    
    # 查找所有result.txt文件
    # pattern = os.path.join(results_path, "**/result.txt")
    # pattern = '/home/agent/code/OSWorld/results/pyautogui/screenshot/cu_sft_0902_nothink_history_steponly_singletask_osworld_1512/*/*/result.txt'
    pattern = f'/home/agent/code/OSWorld/results/pyautogui/screenshot/{model_name}/*/*/result.txt'
    result_files = glob.glob(pattern, recursive=True)
    
    total_tasks = 0
    successful_tasks = 0
    
    for file_path in result_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
            
            score = float(content)
            total_tasks += 1
            
            if score >= 0.8:  # 阈值: 0.8及以上视为成功
                successful_tasks += 1
                
        except (ValueError, IOError):
            # 忽略无法处理的文件
            continue
    
    if total_tasks == 0:
        print("未找到有效的结果文件")
        return
    
    success_rate = successful_tasks / total_tasks * 100
    
    print(f"总任务数: {total_tasks}")
    print(f"成功任务数: {successful_tasks}")
    print(f"失败任务数: {total_tasks - successful_tasks}")
    print(f"成功率: {success_rate:.2f}%")
    
    return {
        'total': total_tasks,
        'success': successful_tasks,
        'rate': success_rate
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='计算OSWorld指定模型的成功率')
    parser.add_argument('--model_name', type=str, required=True, 
                       help='模型名称 (例如: swift_sft_simple_multi_nodes)')
    parser.add_argument('--results_path', type=str, default=None,
                       help='结果文件路径 (默认: /home/agent/code/OSWorld/results)')
    
    args = parser.parse_args()
    
    quick_calculate(args.results_path, args.model_name)
