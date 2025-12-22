#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
改进的输入方法集合 - 解决Docker虚拟机中剪贴板输入问题
"""

import subprocess
import pyautogui
import time
import tempfile
import os
import sys
from typing import Optional, Union


def check_xclip_available() -> bool:
    """检查xclip是否可用"""
    try:
        result = subprocess.run(['which', 'xclip'], capture_output=True, text=True)
        return result.returncode == 0
    except:
        return False


def check_pyperclip_available() -> bool:
    """检查pyperclip是否可用"""
    try:
        import pyperclip
        # 简单测试
        test_str = "test"
        pyperclip.copy(test_str)
        return pyperclip.paste() == test_str
    except:
        return False


def type_with_xclip(text: str, retry_count: int = 2) -> bool:
    """
    使用xclip实现剪贴板输入
    
    Args:
        text: 要输入的文本
        retry_count: 重试次数
    
    Returns:
        bool: 是否成功
    """
    if not check_xclip_available():
        print("❌ xclip不可用")
        return False
    
    for attempt in range(retry_count + 1):
        try:
            # 使用xclip设置剪贴板 (主剪贴板和选择缓冲区都设置)
            for selection in ['clipboard', 'primary']:
                process = subprocess.Popen(
                    ['xclip', '-selection', selection], 
                    stdin=subprocess.PIPE, 
                    text=True,
                    stderr=subprocess.PIPE
                )
                stdout, stderr = process.communicate(input=text)
                
                if process.returncode != 0:
                    print(f"⚠️ xclip设置{selection}失败: {stderr}")
                    continue
            
            # 等待剪贴板更新
            time.sleep(0.2)
            
            # 验证剪贴板内容
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                  capture_output=True, text=True, timeout=2)
            
            if result.returncode == 0 and result.stdout == text:
                # 执行粘贴操作
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.3)
                
                print(f"✅ xclip方法成功 (第{attempt + 1}次尝试)")
                return True
            else:
                print(f"⚠️ 剪贴板验证失败 (第{attempt + 1}次尝试)")
                
        except subprocess.TimeoutExpired:
            print(f"⚠️ xclip超时 (第{attempt + 1}次尝试)")
        except Exception as e:
            print(f"⚠️ xclip方法失败 (第{attempt + 1}次尝试): {e}")
        
        if attempt < retry_count:
            time.sleep(0.5)  # 重试前等待
    
    print("❌ xclip方法所有尝试都失败")
    return False


def type_with_pyperclip(text: str, retry_count: int = 2) -> bool:
    """
    使用pyperclip实现剪贴板输入
    
    Args:
        text: 要输入的文本
        retry_count: 重试次数
    
    Returns:
        bool: 是否成功
    """
    if not check_pyperclip_available():
        print("❌ pyperclip不可用")
        return False
    
    try:
        import pyperclip
        
        for attempt in range(retry_count + 1):
            try:
                # 设置剪贴板
                pyperclip.copy(text)
                time.sleep(0.1)
                
                # 验证剪贴板内容
                clipboard_content = pyperclip.paste()
                if clipboard_content == text:
                    # 执行粘贴
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.3)
                    
                    print(f"✅ pyperclip方法成功 (第{attempt + 1}次尝试)")
                    return True
                else:
                    print(f"⚠️ pyperclip剪贴板验证失败 (第{attempt + 1}次尝试)")
                    
            except Exception as e:
                print(f"⚠️ pyperclip方法失败 (第{attempt + 1}次尝试): {e}")
            
            if attempt < retry_count:
                time.sleep(0.3)
        
        print("❌ pyperclip方法所有尝试都失败")
        return False
        
    except ImportError:
        print("❌ pyperclip模块未安装")
        return False


def type_directly(text: str, interval: float = 0.03, chunk_size: int = 50) -> bool:
    """
    直接使用pyautogui.write()输入文本（分块处理长文本）
    
    Args:
        text: 要输入的文本
        interval: 字符间隔时间
        chunk_size: 分块大小
    
    Returns:
        bool: 是否成功
    """
    try:
        # 对于长文本，分块处理以避免超时或缓冲区问题
        if len(text) > chunk_size:
            for i in range(0, len(text), chunk_size):
                chunk = text[i:i + chunk_size]
                pyautogui.write(chunk, interval=interval)
                time.sleep(0.1)  # 分块间的暂停
        else:
            pyautogui.write(text, interval=interval)
        
        time.sleep(0.2)
        print("✅ 直接输入方法成功")
        return True
        
    except Exception as e:
        print(f"❌ 直接输入方法失败: {e}")
        return False


def type_with_temp_file(text: str) -> bool:
    """
    使用临时文件配合xclip输入
    
    Args:
        text: 要输入的文本
    
    Returns:
        bool: 是否成功
    """
    if not check_xclip_available():
        print("❌ xclip不可用，无法使用临时文件方法")
        return False
    
    temp_path = None
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
            f.write(text)
            temp_path = f.name
        
        # 使用xclip从文件读取到剪贴板
        result = subprocess.run(['xclip', '-selection', 'clipboard', temp_path], 
                              capture_output=True, timeout=5)
        
        if result.returncode == 0:
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.3)
            
            print("✅ 临时文件方法成功")
            return True
        else:
            print("❌ 临时文件方法失败")
            return False
            
    except Exception as e:
        print(f"❌ 临时文件方法失败: {e}")
        return False
    finally:
        # 清理临时文件
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except:
                pass


def type_with_keyboard_simulation(text: str) -> bool:
    """
    使用键盘模拟输入（逐字符）
    
    Args:
        text: 要输入的文本
    
    Returns:
        bool: 是否成功
    """
    try:
        for char in text:
            if char.isalnum() or char in ' .,!?':  # 只处理安全字符
                pyautogui.press(char)
            else:
                # 对于特殊字符，使用write
                pyautogui.write(char, interval=0.05)
            time.sleep(0.02)
        
        print("✅ 键盘模拟方法成功")
        return True
        
    except Exception as e:
        print(f"❌ 键盘模拟方法失败: {e}")
        return False


def type_with_hybrid_method(text: str, prefer_clipboard: bool = True) -> bool:
    """
    混合方法：尝试多种输入方式，按优先级顺序
    
    Args:
        text: 要输入的文本
        prefer_clipboard: 是否优先使用剪贴板方法
    
    Returns:
        bool: 是否成功
    """
    print(f"🔄 开始混合输入方法 (文本长度: {len(text)})")
    
    methods = []
    
    if prefer_clipboard:
        # 优先使用剪贴板方法
        methods = [
            ("pyperclip", lambda: type_with_pyperclip(text)),
            ("xclip", lambda: type_with_xclip(text)),
            ("temp_file", lambda: type_with_temp_file(text)),
            ("direct", lambda: type_directly(text)),
        ]
    else:
        # 优先使用直接输入
        methods = [
            ("direct", lambda: type_directly(text)),
            ("pyperclip", lambda: type_with_pyperclip(text)),
            ("xclip", lambda: type_with_xclip(text)),
            ("temp_file", lambda: type_with_temp_file(text)),
        ]
    
    # 对于长文本或包含特殊字符的文本，优先使用剪贴板
    if len(text) > 100 or any(ord(c) > 127 for c in text):
        prefer_clipboard = True
        print("🔄 检测到长文本或特殊字符，优先使用剪贴板方法")
    
    for method_name, method_func in methods:
        try:
            print(f"🔄 尝试{method_name}方法...")
            if method_func():
                print(f"✅ {method_name}方法成功")
                return True
            time.sleep(0.3)  # 方法间等待
        except Exception as e:
            print(f"❌ {method_name}方法异常: {e}")
            continue
    
    print("❌ 所有输入方法都失败了")
    return False


def smart_type(text: str, **kwargs) -> bool:
    """
    智能输入函数 - 根据文本特征选择最佳输入方法
    
    Args:
        text: 要输入的文本
        **kwargs: 其他参数
    
    Returns:
        bool: 是否成功
    """
    if not text:
        print("⚠️ 输入文本为空")
        return True
    
    # 根据文本特征选择策略
    has_unicode = any(ord(c) > 127 for c in text)
    is_long = len(text) > 50
    has_special_chars = any(c in text for c in ['\n', '\t', '\r'])
    
    print(f"📝 智能输入分析:")
    print(f"   文本长度: {len(text)}")
    print(f"   包含Unicode: {has_unicode}")
    print(f"   包含特殊字符: {has_special_chars}")
    
    # 选择最佳策略
    if has_unicode or is_long or has_special_chars:
        print("🎯 选择剪贴板优先策略")
        return type_with_hybrid_method(text, prefer_clipboard=True)
    else:
        print("🎯 选择直接输入优先策略")
        return type_with_hybrid_method(text, prefer_clipboard=False)


# 兼容性函数 - 替代原有的pyperclip.copy + pyautogui.hotkey方案
def improved_clipboard_input(text: str) -> bool:
    """
    改进的剪贴板输入函数 - 直接替代原有方案
    
    Args:
        text: 要输入的文本
    
    Returns:
        bool: 是否成功
    """
    return smart_type(text)


# 测试函数
def test_all_methods():
    """测试所有输入方法"""
    test_texts = [
        "Korean",
        "Hello World",
        "测试中文输入",
        "Mixed text 混合文本 🚀",
        "Long text with multiple lines\nSecond line\nThird line with special chars: !@#$%",
    ]
    
    print("🧪 开始测试所有输入方法")
    print("=" * 60)
    
    for i, text in enumerate(test_texts, 1):
        print(f"\n测试 {i}: '{text[:30]}{'...' if len(text) > 30 else ''}'")
        print("-" * 40)
        
        success = smart_type(text)
        print(f"结果: {'✅ 成功' if success else '❌ 失败'}")
        
        if i < len(test_texts):
            print("等待3秒后进行下一个测试...")
            time.sleep(3)
    
    print("\n🏁 所有测试完成")


if __name__ == "__main__":
    # 如果直接运行此脚本，执行测试
    if len(sys.argv) > 1:
        test_text = ' '.join(sys.argv[1:])
        print(f"测试输入: {test_text}")
        success = smart_type(test_text)
        print(f"结果: {'成功' if success else '失败'}")
    else:
        test_all_methods()
