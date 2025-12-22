#!/usr/bin/env python3
"""
Docker GUI访问演示

这个脚本演示如何在Docker环境中访问虚拟机的GUI界面
"""

import os
import time
import subprocess
import requests

def check_vnc_access(port):
    """检查VNC端口是否可访问"""
    try:
        # 简单的TCP连接测试
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(3)
        result = sock.connect_ex(('localhost', port))
        sock.close()
        return result == 0
    except:
        return False

def check_web_access(port):
    """检查Web端口是否可访问"""
    try:
        response = requests.get(f'http://localhost:{port}', timeout=3)
        return response.status_code == 200
    except:
        return False

def main():
    print("🐳 Docker GUI访问演示")
    print("=" * 50)
    
    # 运行一个简单的Docker容器来演示
    print("1. 启动Docker环境...")
    
    # 检查Docker是否运行
    try:
        result = subprocess.run(['docker', 'ps'], capture_output=True, text=True)
        if result.returncode != 0:
            print("❌ Docker服务未运行，请先启动Docker")
            return
        print("✅ Docker服务正在运行")
    except FileNotFoundError:
        print("❌ Docker未安装")
        return
    
    print("\n2. Docker GUI访问方式:")
    print("   Docker环境中的虚拟机GUI通过以下方式访问：")
    print("   - VNC端口：用于图形界面连接")
    print("   - Web端口：通过浏览器访问")
    print("   - Server端口：OSWorld服务端口")
    
    print("\n3. 常见端口说明:")
    common_ports = [
        (8006, "VNC端口", "vncviewer localhost:8006"),
        (8008, "VNC端口", "vncviewer localhost:8008"),  
        (5000, "Web服务", "http://localhost:5000"),
        (5002, "Web服务", "http://localhost:5002"),
        (5910, "noVNC", "http://localhost:5910"),
    ]
    
    for port, desc, access in common_ports:
        print(f"   {port:4d} - {desc:10s} - {access}")
    
    print("\n4. 检查可用端口:")
    for port, desc, access in common_ports:
        vnc_available = check_vnc_access(port)
        web_available = check_web_access(port)
        
        if vnc_available or web_available:
            status = "✅ 可用"
            print(f"   端口 {port} ({desc}): {status}")
            print(f"     访问方式: {access}")
        else:
            print(f"   端口 {port} ({desc}): ❌ 不可用")
    
    print("\n💡 使用建议:")
    print("1. 对于Docker环境，GUI访问通过VNC端口而不是直接窗口")
    print("2. 如果要看到执行过程，请：")
    print("   - 安装VNC客户端：sudo apt install vncviewer")
    print("   - 或使用浏览器访问noVNC端口")
    print("3. Docker的 --enable_gui 参数主要是确保VNC服务启动")
    print("4. 实际的可视化需要连接到相应的VNC端口")
    
    print("\n🔧 与传统虚拟化的对比:")
    print("- VirtualBox/VMware: 直接显示虚拟机窗口")
    print("- Docker: 通过VNC端口访问虚拟机界面") 
    print("- AWS/云环境: 通过远程桌面或VNC访问")
    
    print("\n🎯 推荐方案:")
    print("如果您想要直接看到虚拟机窗口，建议使用：")
    print("python3 run_multienv_stepcloud.py --provider_name virtualbox --enable_gui")

if __name__ == "__main__":
    main()



