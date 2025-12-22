from __future__ import annotations
import argparse
import datetime
import json
import logging
import os
import sys
import time
from desktop_env.desktop_env import DesktopEnv

# 加载环境变量
if os.path.exists(".env"):
    from dotenv import load_dotenv
    load_dotenv()

# 配置日志
def setup_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # 设置日志格式
    formatter = logging.Formatter(
        fmt="\x1b[1;33m[%(asctime)s \x1b[31m%(levelname)s \x1b[32m%(module)s/%(lineno)d\x1b[1;33m] \x1b[0m%(message)s"
    )
    console_handler.setFormatter(formatter)
    
    # 添加处理器
    logger.addHandler(console_handler)
    return logger

logger = setup_logger()

def config() -> argparse.Namespace:
    """简化的配置函数，只保留Docker相关参数"""
    parser = argparse.ArgumentParser(
        description="启动Docker环境并显示端口信息"
    )
    
    # Docker相关配置
    parser.add_argument(
        "--path_to_vm", type=str, default=None, help="虚拟机路径"
    )
    parser.add_argument(
        "--provider_name", type=str, default="docker", help="Provider name (固定为docker)"
    )
    parser.add_argument(
        "--headless", action="store_true", help="无头模式运行"
    )
    parser.add_argument(
        "--enable_gui", action="store_true", help="启用GUI显示"
    )
    parser.add_argument(
        "--screen_width", type=int, default=1920, help="屏幕宽度"
    )
    parser.add_argument(
        "--screen_height", type=int, default=1080, help="屏幕高度"
    )
    parser.add_argument(
        "--client_password", type=str, default="", help="客户端密码"
    )
    parser.add_argument(
        "--action_space", type=str, default="pyautogui", help="操作空间类型"
    )
    parser.add_argument(
        "--observation_type", type=str, default="screenshot", help="观察类型"
    )
    parser.add_argument(
        "--task_config", type=str, default=None, help="任务配置JSON文件路径（可选）"
    )
    
    args = parser.parse_args()
    return args


def display_docker_ports_info():
    """显示Docker端口信息"""
    print("=" * 60)
    print("🐳 Docker端口信息 / Docker Ports Information")
    print("=" * 60)
    print("当Docker环境启动后，您可以通过以下端口访问:")
    print("After Docker environment starts, you can access via these ports:")
    print()
    print("🖥️  VNC端口 / VNC Port:")
    print("   - 通常为 8006-8010 之间的端口")
    print("   - 使用命令: vncviewer localhost:<port>")
    print("   - 或安装 xtightvncviewer: sudo apt install xtightvncviewer")
    print()
    print("🌐 Web服务端口 / Web Service Port:")
    print("   - 通常为 5000-5010 之间的端口")
    print("   - 访问: http://localhost:<port>")
    print()
    print("🎬 实时页面端口 / Real-time Page Port:")
    print("   - 通常为 5000-5010 之间的端口 (与Web服务相同)")
    print("   - 可以在浏览器中查看虚拟机画面")
    print()
    print("📱 Chrome调试端口 / Chrome Debug Port:")
    print("   - 通常为 9222-9230 之间的端口")
    print("   - 访问: http://localhost:<port>")
    print("=" * 60)
    print()


def get_docker_container_info():
    """获取Docker容器详细信息"""
    containers = []
    
    # 首先尝试获取Docker容器信息
    try:
        import subprocess
        import json
        
        # 尝试使用sudo权限获取Docker容器信息
        result = subprocess.run(['sudo', 'docker', 'ps', '--format', 'json'], 
                              capture_output=True, text=True, check=True)
        
        for line in result.stdout.strip().split('\n'):
            if line:
                try:
                    container_info = json.loads(line)
                    containers.append(container_info)
                except json.JSONDecodeError:
                    continue
    except Exception as e:
        logger.debug(f"Docker命令失败: {e}")
    
    # 如果Docker容器为空，尝试获取QEMU进程信息
    if not containers:
        try:
            result = subprocess.run(['ps', 'aux'], capture_output=True, text=True, check=True)
            for line in result.stdout.split('\n'):
                # 查找QEMU进程，包含ubuntu或kvm关键词
                if ('qemu' in line.lower() and 
                    ('ubuntu' in line.lower() or 'kvm' in line.lower() or 'system' in line.lower())):
                    # 提取进程信息
                    parts = line.split()
                    if len(parts) >= 11:
                        pid = parts[1]
                        cpu = parts[2]
                        mem = parts[3]
                        command = ' '.join(parts[10:])
                        
                        containers.append({
                            'type': 'QEMU虚拟机',
                            'pid': pid,
                            'cpu': cpu,
                            'mem': mem,
                            'command': command,
                            'status': '运行中'
                        })
        except Exception as e2:
            logger.warning(f"无法获取QEMU进程信息: {e2}")
    
    return containers


def display_docker_startup_details(env):
    """显示Docker启动详细信息"""
    print("=" * 60)
    print("🚀 Docker启动详细信息 / Docker Startup Details")
    print("=" * 60)
    
    # 显示容器/虚拟机信息
    containers = get_docker_container_info()
    if containers:
        print(f"📦 运行中的容器/虚拟机数量: {len(containers)}")
        for i, container in enumerate(containers, 1):
            if 'type' in container:
                # QEMU虚拟机信息
                print(f"   虚拟机 {i}: {container.get('type', 'Unknown')}")
                print(f"   进程ID: {container.get('pid', 'Unknown')}")
                print(f"   CPU使用: {container.get('cpu', 'Unknown')}%")
                print(f"   内存使用: {container.get('mem', 'Unknown')}%")
                print(f"   状态: {container.get('status', 'Unknown')}")
                print(f"   命令: {container.get('command', 'Unknown')[:150]}...")
            else:
                # Docker容器信息
                print(f"   容器 {i}: {container.get('Names', 'Unknown')} (ID: {container.get('ID', 'Unknown')[:12]})")
                print(f"   状态: {container.get('Status', 'Unknown')}")
                print(f"   端口映射: {container.get('Ports', 'None')}")
            print()
    else:
        print("⚠️  未找到运行中的Docker容器或虚拟机")
    
    # 显示环境配置信息
    if hasattr(env, 'controller') and env.controller:
        print("🔧 环境控制器信息:")
        print(f"   类型: {type(env.controller).__name__}")
        if hasattr(env.controller, 'container_name'):
            print(f"   容器名称: {env.controller.container_name}")
        if hasattr(env.controller, 'ports'):
            print(f"   端口配置: {env.controller.ports}")
        if hasattr(env.controller, 'vm_process'):
            print(f"   虚拟机进程: {env.controller.vm_process}")
    
    print("=" * 60)
    print()


def load_task_config(config_path):
    """加载任务配置文件"""
    if not config_path:
        return None
    
    if not os.path.exists(config_path):
        logger.error(f"❌ 任务配置文件不存在: {config_path}")
        return None
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        logger.info(f"✅ 成功加载任务配置: {config_path}")
        if "instruction" in config:
            logger.info(f"📝 任务说明: {config['instruction']}")
        if "snapshot" in config:
            logger.info(f"📸 快照类型: {config['snapshot']}")
        return config
    except Exception as e:
        logger.error(f"❌ 加载任务配置文件失败: {e}")
        return None


def setup_initial_state(env, task_config):
    """根据任务配置设置初始状态"""
    if not task_config or "config" not in task_config:
        logger.info("📋 使用默认初始状态（桌面环境）")
        return True
    
    logger.info("⚙️ 正在根据任务配置设置初始状态...")
    
    try:
        # 直接使用 setup_controller 执行配置，而不是通过reset
        config_list = task_config["config"]
        logger.info(f"📋 执行 {len(config_list)} 个配置步骤...")
        
        success = env.setup_controller.setup(config_list, use_proxy=False)
        
        if success:
            logger.info("✅ 初始状态设置完成!")
            return True
        else:
            logger.error("❌ 配置执行失败")
            return False
            
    except Exception as e:
        logger.error(f"❌ 设置初始状态失败: {e}")
        import traceback
        logger.error(f"详细错误信息: {traceback.format_exc()}")
        return False


def start_docker_environment(args):
    """启动Docker环境"""
    logger.info("🚀 正在启动Docker环境...")
    
    try:
        # 创建Docker环境
        env = DesktopEnv(
            path_to_vm=args.path_to_vm,
            action_space=args.action_space,
            provider_name=args.provider_name,
            headless=args.headless and not args.enable_gui,
            os_type="Ubuntu",
            require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
            enable_proxy=False,  # 简化配置，不使用代理
            client_password=args.client_password
        )
        
        logger.info("✅ Docker环境启动成功!")
        
        # 等待环境完全初始化
        logger.info("⏳ 等待环境完全初始化...")
        time.sleep(15)  # 增加等待时间确保Docker容器完全启动
        
        # 加载任务配置并设置初始状态
        task_config = load_task_config(args.task_config)
        if not setup_initial_state(env, task_config):
            logger.warning("⚠️ 初始状态设置失败，但环境仍可使用")
        
        return env
        
    except Exception as e:
        logger.error(f"❌ Docker环境启动失败: {e}")
        raise


def monitor_docker_environment(env, duration=60):
    """监控Docker环境运行状态"""
    logger.info(f"🔍 开始监控Docker环境，持续 {duration} 秒...")
    
    start_time = time.time()
    while time.time() - start_time < duration:
        try:
            # 显示当前时间
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            logger.info(f"⏰ 当前时间: {current_time}")
            
            # 显示Docker容器状态
            containers = get_docker_container_info()
            if containers:
                logger.info(f"📦 运行中的容器: {len(containers)}")
                for container in containers:
                    logger.info(f"   - {container.get('Names', 'Unknown')}: {container.get('Status', 'Unknown')}")
            else:
                logger.warning("⚠️  未检测到运行中的Docker容器")
            
            # 等待5秒后继续监控
            time.sleep(5)
            
        except KeyboardInterrupt:
            logger.info("🛑 用户中断监控")
            break
        except Exception as e:
            logger.error(f"❌ 监控过程中出现错误: {e}")
            break
    
    logger.info("🏁 监控结束")


if __name__ == "__main__":
    """
    主函数：启动Docker环境并显示详细信息
    
    使用方法 / Usage:
    
    1. 启动默认桌面环境:
       python run_docker.py
       
    2. 启动指定任务环境:
       python run_docker.py --task_config evaluation_examples/examples/chrome/example.json
       
    2a. 测试简单任务配置:
       python run_docker.py --task_config simple_task_config.json
       
    3. 启用GUI显示:
       python run_docker.py --enable_gui
       
    4. 无头模式运行:
       python run_docker.py --headless
       
    5. 指定屏幕分辨率:
       python run_docker.py --screen_width 1920 --screen_height 1080
       
    任务配置文件格式 / Task Config Format:
    {
        "id": "task_id",
        "snapshot": "chrome|gimp|os|multiapps|libreoffice_calc|libreoffice_writer",
        "instruction": "任务说明",
        "config": [
            {
                "type": "launch|download|execute|open",
                "parameters": {...}
            }
        ]
    }
    
    如果不指定 --task_config 参数，将使用默认的桌面初始状态。
    """
    print("=" * 60)
    print("🐳 Docker环境启动器 / Docker Environment Launcher")
    print("=" * 60)
    print("此工具用于启动Docker环境并显示相关端口信息")
    print("This tool is used to start Docker environment and display port information")
    print("=" * 60)
    print()
    
    try:
        # 解析命令行参数
        args = config()
        
        # 显示Docker端口信息
        display_docker_ports_info()
        
        # 显示配置信息
        logger.info("=" * 60)
        logger.info("🔧 Docker配置信息 / Docker Configuration:")
        logger.info(f"  Provider: {args.provider_name}")
        logger.info(f"  Headless Mode: {args.headless}")
        logger.info(f"  GUI Display: {args.enable_gui}")
        logger.info(f"  Screen Size: {args.screen_width}x{args.screen_height}")
        logger.info(f"  Action Space: {args.action_space}")
        logger.info(f"  Observation Type: {args.observation_type}")
        if args.task_config:
            logger.info(f"  Task Config: {args.task_config}")
        else:
            logger.info(f"  Task Config: 使用默认桌面环境")
        logger.info("=" * 60)
        print()
        
        # 启动Docker环境
        env = start_docker_environment(args)
        
        # 显示Docker启动详细信息
        display_docker_startup_details(env)
        
        # 监控Docker环境（可选）
        try:
            logger.info("按 Ctrl+C 停止监控...")
            monitor_docker_environment(env, duration=300)  # 监控5分钟
        except KeyboardInterrupt:
            logger.info("🛑 用户停止监控")
        
        # 关闭环境
        logger.info("🔄 正在关闭Docker环境...")
        try:
            env.close()
            logger.info("✅ Docker环境已成功关闭")
        except Exception as e:
            logger.error(f"❌ 关闭Docker环境时出现错误: {e}")
            
    except KeyboardInterrupt:
        logger.info("🛑 用户中断程序")
    except Exception as e:
        logger.error(f"❌ 程序执行过程中出现错误: {e}")
        import traceback
        logger.error(traceback.format_exc())
    finally:
        logger.info("🏁 程序结束")
