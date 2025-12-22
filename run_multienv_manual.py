from __future__ import annotations
import argparse
import datetime
import json
import logging
import os
import sys
import signal
import time
from typing import List, Dict
import math
from tqdm import tqdm
from multiprocessing import Process, Manager, Queue
from multiprocessing import current_process
import lib_run_single
from desktop_env.desktop_env import DesktopEnv

# Global variables for signal handling
active_environments = []
processes = []
is_terminating = False

# import wandb

# load the environment variables from .env file
if os.path.exists(".env"):
    from dotenv import load_dotenv
    load_dotenv()

#  Logger Configs {{{ #
def config() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run end-to-end evaluation on the benchmark"
    )

    # environment config
    parser.add_argument("--path_to_vm", type=str, default=None)
    parser.add_argument(
        "--headless", action="store_true", help="Run in headless machine (disable to see VM execution process)"
    )
    parser.add_argument(
        "--enable_gui", action="store_true", help="Enable GUI display for VM execution process"
    )
    parser.add_argument(
        "--action_space", type=str, default="pyautogui", help="Action type"
    )
    parser.add_argument(
        "--observation_type",
        choices=["screenshot", "a11y_tree", "screenshot_a11y_tree", "som"],
        default="screenshot",
        help="Observation type",
    )
    parser.add_argument("--sleep_after_execution", type=float, default=0.0)
    parser.add_argument("--max_steps", type=int, default=15)

    # agent config
    parser.add_argument("--max_trajectory_length", type=int, default=3)
    parser.add_argument(
        "--test_config_base_dir", type=str, default="evaluation_examples"
    )

    # lm config
    parser.add_argument("--model", type=str, default="ovr_zero_dogegg_g5_fixres_mount_rftdata_it500_out")
    parser.add_argument("--temperature", type=float, default=0.2)
    parser.add_argument("--top_p", type=float, default=0.9)
    parser.add_argument("--max_tokens", type=int, default=8192)
    parser.add_argument("--stop_token", type=str, default=None)
    parser.add_argument("--add_thought_prefix", action="store_true", help="Add thought prefix to the response")
    
    # example config
    parser.add_argument("--domain", type=str, default="all")
    parser.add_argument(
        "--test_all_meta_path", type=str, default="evaluation_examples/test_all.json"
    )
    parser.add_argument(
        "--task_file", type=str, default=None, 
        help="指定单个任务文件进行测试，例如: evaluation_examples/examples/chrome/change_chrome_download_path.json"
    )

    # logging related
    parser.add_argument("--result_dir", type=str, default="./results")
    parser.add_argument("--num_envs", type=int, default=1, help="Number of environments to run in parallel")  
    parser.add_argument("--log_level", type=str, choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'], 
                       default='INFO', help="Set the logging level")
    parser.add_argument("--max_tasks", type=int, default=None, help="Maximum number of tasks to execute (None for all tasks)")
    # aws config
    parser.add_argument(
        "--region", type=str, default="us-east-1", help="AWS region for the VM"
    )
    parser.add_argument(
        "--provider_name", type=str, default="aws", choices=["aws", "virtualbox", "vmware", "docker", "azure"], help="Provider name"
    )
    parser.add_argument(
        "--client_password", type=str, default="", help="Client password"
    )
    parser.add_argument(
        "--screen_width", type=int, default=1920, help="Screen width"
    )
    parser.add_argument(
        "--screen_height", type=int, default=1080, help="Screen height"
    )
    parser.add_argument(
        "--enable_network", action="store_true", default=True, help="Enable network connection for VM"
    )
    parser.add_argument(
        "--nat_network", action="store_true", default=True, help="Enable NAT network for VM internet access"
    )
    parser.add_argument(
        "--bridge_network", action="store_true", help="Enable bridge network for VM (allows direct network access)"
    )
    parser.add_argument(
        "--proxy_ip", type=str, default=None, help="Proxy server IP address for Docker containers"
    )
    parser.add_argument(
        "--proxy_port", type=int, default=8080, help="Proxy server port (default: 8080)"
    )
    parser.add_argument(
        "--proxy_username", type=str, default=None, help="Proxy username (if authentication required)"
    )
    parser.add_argument(
        "--proxy_password", type=str, default=None, help="Proxy password (if authentication required)"
    )
    args = parser.parse_args()
    return args

args = config()  # Get command line arguments first

logger = logging.getLogger()
log_level = getattr(logging, args.log_level.upper())
logger.setLevel(log_level)

datetime_str: str = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")

# 创建logs目录
os.makedirs("logs", exist_ok=True)

# 定义日志文件路径
log_dir = os.path.abspath("logs")
normal_log_path = os.path.join(log_dir, "normal-{:}.log".format(datetime_str))
debug_log_path = os.path.join(log_dir, "debug-{:}.log".format(datetime_str))

file_handler = logging.FileHandler(normal_log_path, encoding="utf-8")
debug_handler = logging.FileHandler(debug_log_path, encoding="utf-8")
stdout_handler = logging.StreamHandler(sys.stdout)

file_handler.setLevel(logging.INFO)
debug_handler.setLevel(logging.DEBUG)
stdout_handler.setLevel(log_level)

formatter = logging.Formatter(
    fmt="\x1b[1;33m[%(asctime)s \x1b[31m%(levelname)s \x1b[32m%(module)s/%(lineno)d-%(processName)s\x1b[1;33m] \x1b[0m%(message)s"
)
file_handler.setFormatter(formatter)
debug_handler.setFormatter(formatter)
stdout_handler.setFormatter(formatter)

stdout_handler.addFilter(logging.Filter("desktopenv"))

logger.addHandler(file_handler)
logger.addHandler(debug_handler)
logger.addHandler(stdout_handler)
#  }}} Logger Configs #

logger = logging.getLogger("desktopenv.experiment")


def distribute_tasks(test_all_meta: dict, max_tasks: int = None) -> List[tuple]:
    all_tasks = []
    for domain, examples in test_all_meta.items():
        for example_id in examples:
            all_tasks.append((domain, example_id))
    
    # 限制任务数量
    if max_tasks is not None and max_tasks > 0:
        all_tasks = all_tasks[:max_tasks]
        logger.info(f"限制执行任务数量为: {max_tasks}")
    
    return all_tasks


def display_log_info():
    """显示日志存储信息"""
    print("=" * 60)
    print("📝 日志信息 / Log Information")
    print("=" * 60)
    print(f"📁 日志目录 / Log Directory: {log_dir}")
    print(f"📄 常规日志 / Normal Log: {normal_log_path}")
    print(f"🔍 调试日志 / Debug Log: {debug_log_path}")
    print("=" * 60)
    print()


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


def test_network_connectivity(env):
    """测试虚拟机网络连接"""
    try:
        logger.info("Testing network connectivity...")
        # 对于Docker环境，网络连接通常是自动的
        logger.info("Network connectivity assumed available for containerized environment")
        return True
    except Exception as e:
        logger.error(f"Network connectivity test error: {e}")
        return False


def get_host_ip_for_docker():
    """获取宿主机在Docker网络中的IP地址"""
    try:
        import subprocess
        
        # 方法1: 通过Docker网关IP获取宿主机IP
        result = subprocess.run(['docker', 'network', 'inspect', 'bridge'], 
                              capture_output=True, text=True, check=True)
        import json
        network_info = json.loads(result.stdout)
        gateway_ip = network_info[0]['IPAM']['Config'][0]['Gateway']
        logger.info(f"Docker网关IP: {gateway_ip}")
        return gateway_ip
    except Exception as e:
        logger.warning(f"无法获取Docker网关IP: {e}")
        
        # 方法2: 使用默认Docker网关
        return "172.17.0.1"


def setup_docker_proxy(env, args):
    """设置Docker容器内的网络代理 - 已禁用以提高网络稳定性"""
    logger.info("代理配置已禁用，使用直接网络连接以提高稳定性")
    return True


def setup_vm_network(env, args):
    """设置虚拟机网络配置"""
    try:
        logger.info("Setting up VM network configuration...")
        
        # 对于Docker环境，配置代理
        if args.provider_name == "docker":
            logger.info("Docker environment detected - configuring proxy if specified")
            
            # 设置Docker代理
            proxy_success = setup_docker_proxy(env, args)
            if not proxy_success:
                logger.warning("Proxy setup failed, but continuing...")
            
            # 测试网络连接
            return test_network_connectivity(env)
        
        # 对于其他虚拟化环境，使用setup_controller进行网络配置
        if hasattr(env, 'setup_controller') and env.setup_controller:
            # 检查当前网络状态
            try:
                network_status = env.setup_controller._execute_setup("ip addr show")
                logger.info(f"Current network interfaces available")
            except Exception as e:
                logger.warning(f"Could not check network interfaces: {e}")
            
            # 尝试配置网络连接
            if args.enable_network:
                logger.info("Configuring network connection...")
                
                # 尝试启动网络管理器
                try:
                    env.setup_controller._execute_setup("sudo systemctl start NetworkManager")
                    logger.info("NetworkManager started")
                except Exception as e:
                    logger.warning(f"Could not start NetworkManager: {e}")
                
                # 如果启用了NAT网络，尝试配置DHCP
                if args.nat_network:
                    logger.info("Configuring DHCP for network interfaces...")
                    try:
                        # 尝试为所有网络接口配置DHCP
                        env.setup_controller._execute_setup("sudo dhclient")
                        logger.info("DHCP configuration attempted")
                    except Exception as e:
                        logger.warning(f"DHCP configuration failed: {e}")
                
                # 设置DNS服务器
                try:
                    env.setup_controller._execute_setup("echo 'nameserver 8.8.8.8\nnameserver 8.8.4.4' | sudo tee /etc/resolv.conf")
                    logger.info("DNS servers configured (8.8.8.8, 8.8.4.4)")
                except Exception as e:
                    logger.warning(f"DNS setup failed: {e}")
                    
                # 为非Docker环境设置代理
                if args.proxy_ip:
                    proxy_success = setup_docker_proxy(env, args)
                    if not proxy_success:
                        logger.warning("Proxy setup failed, but continuing...")
        else:
            logger.warning("No setup_controller available for network configuration")
        
        # 测试网络连接
        network_ok = test_network_connectivity(env)
        if network_ok:
            logger.info("✅ VM network setup completed successfully")
        else:
            logger.warning("⚠️ VM network setup completed but connectivity test failed")
            
        return network_ok
    except Exception as e:
        logger.error(f"VM network setup error: {e}")
        return False


def process_signal_handler(signum, frame, env_idx):
    """Signal handler for child processes to gracefully shut down their environments."""
    logger.info(f"Process {env_idx + 1} received signal {signum}. Shutting down...")
    
    # Get the active_environments from the caller's frame
    local_vars = frame.f_locals
    active_environments = local_vars.get('active_environments', [])
    
    # Close environment in the current process context
    for env in active_environments:
        if env is not None:
            try:
                logger.info(f"Process {env_idx + 1} closing environment...")
                env.close()
                logger.info(f"Process {env_idx + 1} environment closed successfully")
            except Exception as e:
                logger.error(f"Process {env_idx + 1} error closing environment: {e}")
    
    logger.info(f"Process {env_idx + 1} shutdown complete. Exiting.")
    sys.exit(0)


def run_env_tasks(task_queue: Queue, args: argparse.Namespace, shared_scores: list):
    active_environments = []
    env = None
    try:
        # 根据 provider_name 动态导入相应的模块
        if args.provider_name == "aws":
            from desktop_env.providers.aws.manager import IMAGE_ID_MAP
            REGION = args.region
            screen_size = (args.screen_width, args.screen_height)
            ami_id = IMAGE_ID_MAP[REGION].get(screen_size, IMAGE_ID_MAP[REGION][(1920, 1080)])
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                region=REGION,
                snapshot_name=ami_id,
                screen_size=screen_size,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "docker":
            # Docker 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "virtualbox":
            # VirtualBox 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "vmware":
            # VMware 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        else:
            # 默认使用本地环境
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        active_environments.append(env)
        
        # 配置虚拟机网络（针对不同的提供商）
        if args.enable_network:
            logger.info("Configuring VM network...")
            try:
                if args.provider_name == "virtualbox" and args.bridge_network:
                    logger.info("Configuring VirtualBox bridge network...")
                    # 注意：这里只是展示逻辑，实际的VirtualBox网络配置需要在VM启动前完成
                    logger.info("Bridge network should be configured in VirtualBox settings")
                
                # 通用网络设置
                network_success = setup_vm_network(env, args)
                if not network_success:
                    logger.warning("Network setup failed, but continuing execution...")
                else:
                    logger.info("✅ Network configuration completed successfully")
            except Exception as e:
                logger.error(f"Network configuration error: {e}")
                logger.warning("Continuing execution without network configuration...")
        
        # 显示GUI相关信息
        if args.enable_gui and not args.headless:
            logger.info("🖥️ VM GUI display is enabled - you should see the VM window")
            logger.info("VM will be visible during execution for monitoring purposes")
        elif args.headless:
            logger.info("VM is running in headless mode - no GUI window will be shown")
        
        # 人工操作模式，不需要AI代理
        agent = None
        logger.info(f"Process {current_process().name} started.")
        while True:
            try:
                item = task_queue.get(timeout=5)
            except Exception:
                break
            domain, example_id = item
            try:
                config_file = os.path.join(
                    args.test_config_base_dir, f"examples/{domain}/{example_id}.json"
                )
                with open(config_file, "r", encoding="utf-8") as f:
                    example = json.load(f)
                logger.info(f"[{current_process().name}][Domain]: {domain}")
                logger.info(f"[{current_process().name}][Example ID]: {example_id}")
                logger.info(f"[{current_process().name}][Instruction]: {example['instruction']}")
                example_result_dir = os.path.join(
                    args.result_dir,
                    args.action_space,
                    args.observation_type,
                    args.model,
                    domain,
                    example_id,
                )
                os.makedirs(example_result_dir, exist_ok=True)
                try:
                    # 使用人工操作模式
                    lib_run_single.run_single_example_human(
                        env,
                        example,
                        args.max_steps,
                        example["instruction"],
                        args,
                        example_result_dir,
                        shared_scores,
                    )
                except Exception as e:
                    import traceback
                    logger.error(f"Exception in {current_process().name} {domain}/{example_id}: {e}")
                    logger.error(traceback.format_exc())
                    try:
                        env.controller.end_recording(
                            os.path.join(example_result_dir, "recording.mp4")
                        )
                    except Exception as rec_e:
                        logger.error(f"Failed to end recording: {rec_e}")
                    with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                        f.write(
                            json.dumps(
                                {"Error": f"{domain}/{example_id} - {e}"}
                            )
                        )
                        f.write("\n")
            except Exception as e:
                logger.error(f"Task-level error in {current_process().name}: {e}")
                import traceback
                logger.error(traceback.format_exc())
    except Exception as e:
        logger.error(f"Process-level error in {current_process().name}: {e}")
        import traceback
        logger.error(traceback.format_exc())
    finally:
        logger.info(f"{current_process().name} cleaning up environment...")
        try:
            if env:
                env.close()
                logger.info(f"{current_process().name} environment closed successfully")
        except Exception as e:
            logger.error(f"{current_process().name} error during environment cleanup: {e}")


def signal_handler(signum, frame):
    """Handle termination signals (SIGINT, SIGTERM) to gracefully shutdown environments."""
    global is_terminating, active_environments, processes
    
    # Avoid duplicate handling
    if is_terminating:
        return
    
    is_terminating = True
    logger.info(f"Received signal {signum}. Gracefully shutting down...")
    
    # Close all registered environments in the main process
    for env in active_environments:
        try:
            logger.info(f"Closing environment...")
            env.close()
            logger.info(f"Environment closed successfully")
        except Exception as e:
            logger.error(f"Error closing environment: {e}")
    
    # Send termination signal to all child processes first
    for p in processes:
        if p.is_alive():
            try:
                logger.info(f"Sending termination signal to process {p.name}...")
                p.terminate()
            except Exception as e:
                logger.error(f"Error sending termination signal to process: {e}")
    
    # Allow a short time for processes to handle their own cleanup
    time.sleep(1)
    
    # Forcefully terminate any processes that didn't exit
    for p in processes:
        if p.is_alive():
            try:
                logger.info(f"Forcefully terminating process {p.name}...")
                import signal as sig
                os.kill(p.pid, sig.SIGKILL)
            except Exception as e:
                logger.error(f"Error forcefully terminating process: {e}")
    
    logger.info("Shutdown complete. Exiting.")
    sys.exit(0)


def test_manual_mode(args: argparse.Namespace, test_all_meta: dict) -> None:
    """人工操作模式：单进程顺序执行任务"""
    logger.info("Args: %s", args)
    all_tasks = distribute_tasks(test_all_meta, args.max_tasks)
    logger.info(f"Total tasks: {len(all_tasks)}")
    
    if args.max_tasks is not None:
        logger.info(f"任务限制: 只执行前 {args.max_tasks} 个任务 / Task limit: Only execute first {args.max_tasks} tasks")
    
    print("\n" + "="*80)
    print("🎯 OSWorld 人工操作模式 / OSWorld Manual Operation Mode")
    print("="*80)
    print(f"📊 总任务数 / Total Tasks: {len(all_tasks)}")
    print("🔄 任务将顺序执行，每个任务完成后进行下一个")
    print("🔄 Tasks will be executed sequentially, one after another")
    print("="*80)
    
    scores = []
    
    # 创建单个环境
    env = None
    try:
        # 根据 provider_name 动态导入相应的模块
        if args.provider_name == "aws":
            from desktop_env.providers.aws.manager import IMAGE_ID_MAP
            REGION = args.region
            screen_size = (args.screen_width, args.screen_height)
            ami_id = IMAGE_ID_MAP[REGION].get(screen_size, IMAGE_ID_MAP[REGION][(1920, 1080)])
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                region=REGION,
                snapshot_name=ami_id,
                screen_size=screen_size,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "docker":
            # Docker 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "virtualbox":
            # VirtualBox 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        elif args.provider_name == "vmware":
            # VMware 环境配置
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        else:
            # 默认使用本地环境
            env = DesktopEnv(
                path_to_vm=args.path_to_vm,
                action_space=args.action_space,
                provider_name=args.provider_name,
                headless=args.headless and not args.enable_gui,
                os_type="Ubuntu",
                require_a11y_tree=args.observation_type in ["a11y_tree", "screenshot_a11y_tree", "som"],
                enable_proxy=args.enable_network,
                client_password=args.client_password
            )
        
        # 配置虚拟机网络（针对不同的提供商）
        if args.enable_network:
            logger.info("Configuring VM network...")
            try:
                if args.provider_name == "virtualbox" and args.bridge_network:
                    logger.info("Configuring VirtualBox bridge network...")
                    logger.info("Bridge network should be configured in VirtualBox settings")
                
                # 通用网络设置
                network_success = setup_vm_network(env, args)
                if not network_success:
                    logger.warning("Network setup failed, but continuing execution...")
                else:
                    logger.info("✅ Network configuration completed successfully")
            except Exception as e:
                logger.error(f"Network configuration error: {e}")
                logger.warning("Continuing execution without network configuration...")
        
        # 显示GUI相关信息
        if args.enable_gui and not args.headless:
            logger.info("🖥️ VM GUI display is enabled - you should see the VM window")
            logger.info("VM will be visible during execution for monitoring purposes")
        elif args.headless:
            logger.info("VM is running in headless mode - no GUI window will be shown")
        
        # 顺序执行每个任务
        for task_idx, (domain, example_id) in enumerate(all_tasks, 1):
            try:
                config_file = os.path.join(
                    args.test_config_base_dir, f"examples/{domain}/{example_id}.json"
                )
                with open(config_file, "r", encoding="utf-8") as f:
                    example = json.load(f)
                
                print(f"\n📋 任务进度 / Task Progress: {task_idx}/{len(all_tasks)}")
                logger.info(f"[Domain]: {domain}")
                logger.info(f"[Example ID]: {example_id}")
                logger.info(f"[Instruction]: {example['instruction']}")
                
                example_result_dir = os.path.join(
                    args.result_dir,
                    args.action_space,
                    args.observation_type,
                    args.model,
                    domain,
                    example_id,
                )
                os.makedirs(example_result_dir, exist_ok=True)
                
                try:
                    # 使用人工操作模式
                    lib_run_single.run_single_example_human(
                        env,
                        example,
                        args.max_steps,
                        example["instruction"],
                        args,
                        example_result_dir,
                        scores,
                    )
                    
                    print(f"✅ 任务 {task_idx}/{len(all_tasks)} 完成")
                    
                    # 如果不是最后一个任务，询问是否继续
                    if task_idx < len(all_tasks):
                        print("\n" + "-"*60)
                        print("⏭️  准备执行下一个任务...")
                        print("⏭️  Preparing for next task...")
                        print("🔄 按回车键继续，或输入'q'退出 / Press Enter to continue, or 'q' to quit")
                        print("-"*60)
                        
                        user_input = input().strip().lower()
                        if user_input == 'q' or user_input == 'quit':
                            print("🛑 用户选择退出 / User chose to quit")
                            break
                    
                except Exception as e:
                    import traceback
                    logger.error(f"Exception in task {domain}/{example_id}: {e}")
                    logger.error(traceback.format_exc())
                    try:
                        env.controller.end_recording(
                            os.path.join(example_result_dir, "recording.mp4")
                        )
                    except Exception as rec_e:
                        logger.error(f"Failed to end recording: {rec_e}")
                    with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                        f.write(
                            json.dumps(
                                {"Error": f"{domain}/{example_id} - {e}"}
                            )
                        )
                        f.write("\n")
                    scores.append(0.0)  # 错误情况下记录0分
                    
            except Exception as e:
                logger.error(f"Task-level error: {e}")
                import traceback
                logger.error(traceback.format_exc())
                scores.append(0.0)  # 错误情况下记录0分
                
    except Exception as e:
        logger.error(f"Environment-level error: {e}")
        import traceback
        logger.error(traceback.format_exc())
    finally:
        logger.info("Cleaning up environment...")
        try:
            if env:
                env.close()
                logger.info("Environment closed successfully")
        except Exception as e:
            logger.error(f"Error during environment cleanup: {e}")
    
    # 显示最终结果
    print("\n" + "="*80)
    print("📊 最终结果 / Final Results")
    print("="*80)
    if scores:
        avg_score = sum(scores) / len(scores)
        print(f"📈 平均分数 / Average Score: {avg_score:.2f}")
        print(f"✅ 成功任务 / Successful Tasks: {sum(1 for s in scores if s >= 1.0)}")
        print(f"⚠️  部分完成 / Partially Completed: {sum(1 for s in scores if 0.5 <= s < 1.0)}")
        print(f"❌ 失败任务 / Failed Tasks: {sum(1 for s in scores if s < 0.5)}")
        logger.info(f"Average score: {avg_score}")
    else:
        print("❌ 没有完成任何任务 / No tasks completed")
        logger.info("Average score: 0")
    print("="*80)


def test(args: argparse.Namespace, test_all_meta: dict) -> None:
    """主测试函数：根据环境模式选择执行方式"""
    global processes
    
    # 如果启用了GUI或者是人工操作模式，强制使用单进程
    if args.enable_gui or args.model == "manual_operation" or args.num_envs == 1:
        logger.info("使用单进程人工操作模式 / Using single-process manual operation mode")
        test_manual_mode(args, test_all_meta)
        return
    
    # 原有的多进程逻辑（保留用于自动化模式）
    logger.info("Args: %s", args)
    all_tasks = distribute_tasks(test_all_meta, args.max_tasks)
    logger.info(f"Total tasks: {len(all_tasks)}")
    
    if args.max_tasks is not None:
        logger.info(f"任务限制: 只执行前 {args.max_tasks} 个任务 / Task limit: Only execute first {args.max_tasks} tasks")
    with Manager() as manager:
        shared_scores = manager.list()
        task_queue = manager.Queue()
        for item in all_tasks:
            task_queue.put(item)
        num_envs = args.num_envs
        processes = []
        for i in range(num_envs):
            p = Process(
                target=run_env_tasks,
                args=(task_queue, args, shared_scores),
                name=f"EnvProcess-{i+1}"
            )
            p.daemon = True
            p.start()
            processes.append(p)
            logger.info(f"Started process {p.name} with PID {p.pid}")
        try:
            while True:
                alive_count = 0
                for idx, p in enumerate(processes):
                    if not p.is_alive():
                        logger.warning(f"Process {p.name} died, restarting...")
                        new_p = Process(
                            target=run_env_tasks,
                            args=(task_queue, args, shared_scores),
                            name=f"EnvProcess-Restart-{idx+1}"
                        )
                        new_p.daemon = True
                        new_p.start()
                        processes[idx] = new_p
                        logger.info(f"Restarted process {new_p.name} with PID {new_p.pid}")
                    else:
                        alive_count += 1
                if task_queue.empty():
                    logger.info("All tasks finished.")
                    break
                if alive_count == 0:
                    logger.error("All processes died, exiting.")
                    break
                time.sleep(5)
            for p in processes:
                p.join()
        except KeyboardInterrupt:
            logger.info("Main process received KeyboardInterrupt. Initiating graceful shutdown...")
            raise
        except Exception as e:
            logger.error(f"Unexpected error while waiting for processes: {e}", exc_info=True)
            for p in processes:
                if p.is_alive():
                    try:
                        logger.info(f"Terminating process {p.name} due to error...")
                        p.terminate()
                    except Exception as term_e:
                        logger.error(f"Error terminating process {p.name}: {term_e}")
            raise
        scores = list(shared_scores)
    logger.info(f"Average score: {sum(scores) / len(scores) if scores else 0}")


def get_unfinished(
    action_space, use_model, observation_type, result_dir, total_file_json
):
    target_dir = os.path.join(result_dir, action_space, observation_type, use_model)

    if not os.path.exists(target_dir):
        return total_file_json

    finished = {}
    for domain in os.listdir(target_dir):
        finished[domain] = []
        domain_path = os.path.join(target_dir, domain)
        if os.path.isdir(domain_path):
            for example_id in os.listdir(domain_path):
                if example_id == "onboard":
                    continue
                example_path = os.path.join(domain_path, example_id)
                if os.path.isdir(example_path):
                    if "result.txt" not in os.listdir(example_path):
                        # empty all files under example_id
                        for file in os.listdir(example_path):
                            os.remove(os.path.join(example_path, file))
                    else:
                        finished[domain].append(example_id)

    if not finished:
        return total_file_json

    for domain, examples in finished.items():
        if domain in total_file_json:
            total_file_json[domain] = [
                x for x in total_file_json[domain] if x not in examples
            ]

    return total_file_json


def get_result(action_space, use_model, observation_type, result_dir, total_file_json):
    target_dir = os.path.join(result_dir, action_space, observation_type, use_model)
    if not os.path.exists(target_dir):
        print("New experiment, no result yet.")
        return None

    all_result = []

    for domain in os.listdir(target_dir):
        domain_path = os.path.join(target_dir, domain)
        if os.path.isdir(domain_path):
            for example_id in os.listdir(domain_path):
                example_path = os.path.join(domain_path, example_id)
                if os.path.isdir(example_path):
                    if "result.txt" in os.listdir(example_path):
                        # empty all files under example_id
                        try:
                            all_result.append(
                                float(
                                    open(
                                        os.path.join(example_path, "result.txt"), "r"
                                    ).read()
                                )
                            )
                        except:
                            all_result.append(0.0)

    if not all_result:
        print("New experiment, no result yet.")
        return None
    else:
        print("Current Success Rate:", sum(all_result) / len(all_result) * 100, "%")
        return all_result


if __name__ == "__main__":
    ####### The complete version of the list of examples #######
    os.environ["TOKENIZERS_PARALLELISM"] = "false"
    
    # Register signal handlers for graceful termination
    signal.signal(signal.SIGINT, signal_handler)  # Handle Ctrl+C
    signal.signal(signal.SIGTERM, signal_handler)  # Handle termination signal
    
    try:
        args = config()
        
        # 显示日志和端口信息
        display_log_info()
        if args.provider_name == "docker":
            display_docker_ports_info()
        
        # 显示配置信息
        logger.info("=" * 60)
        logger.info("VM Configuration:")
        logger.info(f"  Provider: {args.provider_name}")
        logger.info(f"  Headless Mode: {args.headless}")
        logger.info(f"  GUI Display: {args.enable_gui}")
        logger.info(f"  Network Enabled: {args.enable_network}")
        logger.info(f"  NAT Network: {args.nat_network}")
        logger.info(f"  Bridge Network: {args.bridge_network}")
        logger.info(f"  Screen Size: {args.screen_width}x{args.screen_height}")
        logger.info(f"  Number of Environments: {args.num_envs}")
        if args.proxy_ip:
            logger.info(f"  Proxy Server: {args.proxy_ip}:{args.proxy_port}")
            if args.proxy_username:
                logger.info(f"  Proxy Authentication: Enabled (User: {args.proxy_username})")
            else:
                logger.info(f"  Proxy Authentication: None")
        else:
            logger.info(f"  Proxy Server: Not configured")
        logger.info("=" * 60)
        
        if args.enable_gui and not args.headless:
            logger.info("🖥️  GUI模式已启用：您将能够看到虚拟机执行过程")
        
        if args.enable_network:
            logger.info("🌐 网络连接已启用：虚拟机将尝试连接到互联网")
        
        # save args to json in result_dir/action_space/observation_type/model/args.json
        path_to_args = os.path.join(
            args.result_dir,
            args.action_space,
            args.observation_type,
            args.model,
            "args.json",
        )
        os.makedirs(os.path.dirname(path_to_args), exist_ok=True)
        with open(path_to_args, "w", encoding="utf-8") as f:
            json.dump(vars(args), f, indent=4)

        # 如果指定了单个任务文件，则只测试该任务
        if args.task_file:
            logger.info(f"🎯 指定单个任务测试模式 / Single Task Test Mode")
            logger.info(f"📄 任务文件 / Task File: {args.task_file}")
            
            # 检查任务文件是否存在
            if not os.path.exists(args.task_file):
                logger.error(f"❌ 任务文件不存在 / Task file not found: {args.task_file}")
                sys.exit(1)
            
            # 从文件路径中提取 domain 和 example_id
            # 例如: evaluation_examples/examples/chrome/change_chrome_download_path.json
            # 提取: domain=chrome, example_id=change_chrome_download_path
            task_path_parts = args.task_file.replace("\\", "/").split("/")
            
            # 找到 "examples" 的位置
            try:
                examples_idx = task_path_parts.index("examples")
                if examples_idx + 2 < len(task_path_parts):
                    domain = task_path_parts[examples_idx + 1]
                    example_id = os.path.splitext(task_path_parts[-1])[0]  # 去掉 .json 后缀
                else:
                    logger.error(f"❌ 无法从路径中提取 domain 和 example_id: {args.task_file}")
                    sys.exit(1)
            except ValueError:
                logger.error(f"❌ 任务文件路径格式不正确，应包含 'examples' 目录: {args.task_file}")
                sys.exit(1)
            
            logger.info(f"📦 Domain: {domain}")
            logger.info(f"🆔 Example ID: {example_id}")
            
            # 创建只包含这个任务的测试列表
            test_file_list = {domain: [example_id]}
            logger.info(f"✅ 单个任务已加载 / Single task loaded")
        else:
            # 原有的逻辑：从 test_all_meta 加载任务列表
            with open(args.test_all_meta_path, "r", encoding="utf-8") as f:
                test_all_meta = json.load(f)

            if args.domain != "all":
                test_all_meta = {args.domain: test_all_meta[args.domain]}

            test_file_list = get_unfinished(
                args.action_space,
                args.model,
                args.observation_type,
                args.result_dir,
                test_all_meta,
            )
            left_info = ""
            for domain in test_file_list:
                left_info += f"{domain}: {len(test_file_list[domain])}\n"
            logger.info(f"Left tasks:\n{left_info}")

        # 只在非单个任务模式下显示统计信息
        if not args.task_file:
            get_result(
                args.action_space,
                args.model,
                args.observation_type,
                args.result_dir,
                test_all_meta,
            )
        
        test(args, test_file_list)
    except KeyboardInterrupt:
        logger.info("Main process received KeyboardInterrupt.")
        # Signal handler will take care of cleanup
    except Exception as e:
        logger.error(f"Unexpected error in main process: {e}", exc_info=True)
        # Also trigger cleanup for unhandled exceptions
        signal_handler(signal.SIGTERM, None)
    finally:
        # Final cleanup in case any environments or processes remain
        logger.info("Main process final cleanup...")
        for env in active_environments:
            if env is not None:
                try:
                    logger.info(f"Closing environment in final cleanup...")
                    env.close()
                    logger.info(f"Environment closed successfully in final cleanup")
                except Exception as e:
                    logger.error(f"Error during final environment cleanup: {e}")
        
        # First try gentle termination
        for p in processes:
            if p is not None and p.is_alive():
                try:
                    logger.info(f"Terminating process {p.name}...")
                    p.terminate()
                except Exception as e:
                    logger.error(f"Error terminating process: {e}")
        
        # Wait a moment for processes to terminate
        time.sleep(1)
        
        # Then force kill if needed
        for p in processes:
            if p is not None and p.is_alive():
                try:
                    logger.info(f"Force killing process {p.name}...")
                    os.kill(p.pid, signal.SIGKILL)
                    logger.info(f"Process {p.name} force killed")
                except Exception as e:
                    logger.error(f"Error force killing process: {e}")
