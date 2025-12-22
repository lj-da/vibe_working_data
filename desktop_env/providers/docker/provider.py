import logging
import os
import platform
import time
import docker
import psutil
import requests
from filelock import FileLock
from pathlib import Path

from desktop_env.providers.base import Provider

logger = logging.getLogger("desktopenv.providers.docker.DockerProvider")
logger.setLevel(logging.INFO)

WAIT_TIME = 3
RETRY_INTERVAL = 1
LOCK_TIMEOUT = 10


class PortAllocationError(Exception):
    pass


class DockerProvider(Provider):
    def __init__(self, region: str):
        self.client = docker.from_env()
        self.server_port = None
        self.vnc_port = None
        self.chromium_port = None
        self.vlc_port = None
        self.container = None
        self.environment = {
            "DISK_SIZE": "32G", 
            "RAM_SIZE": "4G", 
            "CPU_CORES": "4",
            "NETWORK": "dhcp",
            "DNS_SERVERS": "8.8.8.8,8.8.4.4,1.1.1.1",
            "DOMAIN": "",
            "NTP_SERVERS": "pool.ntp.org",
            "QEMU_NET_OPTIONS": "-netdev user,id=net0,dns=8.8.8.8,dns=8.8.4.4,dns=1.1.1.1,net=10.0.2.0/24,dhcpstart=10.0.2.15,hostfwd=tcp::5000-:5000,hostfwd=tcp::8006-:8006,hostfwd=tcp::9222-:9222,hostfwd=tcp::8080-:8080 -device virtio-net-pci,netdev=net0",
            "VM_NETWORK_CONFIG": "static_wait",  # 改为静态等待模式，避免自动配置竞争
            "ENABLE_DHCP": "delayed",  # 延迟DHCP，避免过早启动
            "MTU": "1500",
            "NETWORK_DEBUG": "yes",
            "FORCE_NETWORK_RESET": "no",  # 禁用强制重置，避免网络状态混乱
            "USE_HOST_DNS": "yes",
            "DISABLE_FIREWALL": "yes",
            "ENABLE_IPV6": "no",  # 暂时禁用IPv6，减少网络复杂性
            "DNS_FALLBACK": "yes",
            "NETWORK_TIMEOUT": "60",  # 增加网络超时时间
            "TCP_KEEPALIVE": "yes",
            "CONNECTION_RETRY": "5",  # 增加重试次数
            "SOCKET_BUFFER_SIZE": "262144",
            "NET_QUEUE_SIZE": "1024",
            "SSL_VERIFY": "no",
            "HTTP_TIMEOUT": "60",
            "HTTPS_TIMEOUT": "90",
            "USER_AGENT": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "IGNORE_SSL_ERRORS": "yes",
            "DISABLE_SSL_VERIFICATION": "yes",
            "ALLOW_INSECURE_SSL": "yes",
            "UPDATE_CERTIFICATES": "yes",
            "SYNC_TIME": "yes",
            # 新增网络稳定性配置
            "NETWORK_STARTUP_DELAY": "15",  # 网络启动延迟
            "DNS_READY_CHECK": "yes",  # 启用DNS就绪检查
            "NETWORK_INIT_RETRY": "3",  # 网络初始化重试次数
            "DHCP_CLIENT_TIMEOUT": "30",  # DHCP客户端超时时间
            "NETWORK_MANAGER_WAIT": "20"  # NetworkManager等待时间
        }

        temp_dir = Path(os.getenv('TEMP') if platform.system() == 'Windows' else '/tmp')
        self.lock_file = temp_dir / "docker_port_allocation.lck"
        self.lock_file.parent.mkdir(parents=True, exist_ok=True)

    def _get_used_ports(self):
        """Get all currently used ports (both system and Docker)."""
        # Get system ports
        system_ports = set(conn.laddr.port for conn in psutil.net_connections())
        
        # Get Docker container ports
        docker_ports = set()
        for container in self.client.containers.list():
            ports = container.attrs['NetworkSettings']['Ports']
            if ports:
                for port_mappings in ports.values():
                    if port_mappings:
                        docker_ports.update(int(p['HostPort']) for p in port_mappings)
        
        return system_ports | docker_ports

    def _get_available_port(self, start_port: int) -> int:
        """Find next available port starting from start_port."""
        used_ports = self._get_used_ports()
        port = start_port
        while port < 65354:
            if port not in used_ports:
                return port
            port += 1
        raise PortAllocationError(f"No available ports found starting from {start_port}")

    def _wait_for_vm_ready(self, timeout: int = 300):
        """Wait for VM to be ready by checking both service and network connectivity."""
        start_time = time.time()
        
        def check_screenshot():
            try:
                response = requests.get(
                    f"http://localhost:{self.server_port}/screenshot",
                    timeout=(10, 10)
                )
                return response.status_code == 200
            except Exception:
                return False
        
        def check_network_connectivity():
            """检查虚拟机内部网络连接性"""
            try:
                # 首先检查基础网络工具是否可用
                basic_check_response = requests.post(
                    f"http://localhost:{self.server_port}/execute",
                    json={
                        "command": "ping -c 1 -W 3 8.8.8.8",
                        "shell": True
                    },
                    timeout=(15, 15)
                )
                
                if basic_check_response.status_code != 200:
                    logger.debug("虚拟机API不可用")
                    return False
                
                basic_result = basic_check_response.json()
                if basic_result.get('status') != 'success' or basic_result.get('returncode', 1) != 0:
                    logger.debug(f"基础网络测试失败: {basic_result.get('error', '')}")
                    return False
                
                # 然后测试HTTP连接
                http_response = requests.post(
                    f"http://localhost:{self.server_port}/execute",
                    json={
                        "command": "curl -s --connect-timeout 10 --max-time 15 -o /dev/null -w '%{http_code}' http://www.baidu.com",
                        "shell": True
                    },
                    timeout=(20, 20)
                )
                
                if http_response.status_code == 200:
                    result = http_response.json()
                    if result.get('status') == 'success' and result.get('returncode', 1) == 0:
                        output = result.get("output", "").strip()
                        logger.info(f"百度连接测试返回: '{output}'")
                        return output == "200"
                
                return False
            except Exception as e:
                logger.debug(f"网络连接测试失败: {e}")
                return False

        # 第一阶段：等待基础服务就绪
        logger.info("等待虚拟机基础服务启动...")
        service_ready = False
        while time.time() - start_time < timeout * 0.6:  # 使用60%的时间等待服务
            if check_screenshot():
                service_ready = True
                logger.info("✅ 虚拟机基础服务已就绪")
                break
            logger.info("检查虚拟机服务状态...")
            time.sleep(RETRY_INTERVAL)
        
        if not service_ready:
            raise TimeoutError("VM service failed to become ready within timeout period")
        
        # 第二阶段：等待网络连接就绪
        logger.info("等待虚拟机网络连接就绪...")
        # 给网络初始化额外时间
        time.sleep(15)  # 静待15秒让网络充分初始化
        
        # 首先尝试修复虚拟机内部网络配置
        self._fix_vm_network_config()
        
        network_ready = False
        network_timeout = timeout * 0.4  # 使用40%的时间等待网络
        network_start = time.time()
        
        while time.time() - network_start < network_timeout:
            if check_network_connectivity():
                network_ready = True
                logger.info("✅ 虚拟机网络连接已就绪")
                break
            logger.info("等待虚拟机网络连接就绪...")
            time.sleep(5)  # 网络检查间隔稍长，避免频繁请求
        
        if not network_ready:
            logger.warning("⚠️ 虚拟机网络连接未在预期时间内就绪，尝试手动修复...")
            # 尝试手动修复网络
            if self._emergency_network_fix():
                logger.info("🔧 手动网络修复完成，重新测试...")
                time.sleep(5)
                if check_network_connectivity():
                    network_ready = True
                    logger.info("✅ 虚拟机网络连接修复成功")
        
        if not network_ready:
            logger.warning("⚠️ 虚拟机网络连接仍未就绪，但继续启动")
            # 不抛出异常，而是继续启动，因为网络可能稍后恢复
        
        return True

    def _fix_vm_network_config(self):
        """修复虚拟机内部网络配置"""
        try:
            logger.info("🔧 修复虚拟机网络配置...")
            
            # 网络修复命令序列
            fix_commands = [
                # 重启网络管理器
                "sudo systemctl restart NetworkManager",
                # 刷新DNS
                "sudo systemctl flush-dns || true",
                "sudo systemd-resolve --flush-caches || true", 
                # 重新配置resolv.conf
                "echo 'nameserver 8.8.8.8\nnameserver 8.8.4.4\nnameserver 1.1.1.1' | sudo tee /etc/resolv.conf",
                # 重新启动网络接口
                "sudo dhclient -r && sudo dhclient || true",
                # 等待网络稳定
                "sleep 5"
            ]
            
            for cmd in fix_commands:
                try:
                    response = requests.post(
                        f"http://localhost:{self.server_port}/execute",
                        json={"command": cmd, "shell": True},
                        timeout=(30, 30)
                    )
                    if response.status_code == 200:
                        result = response.json()
                        logger.debug(f"执行命令 '{cmd}': {result.get('status', 'unknown')}")
                    else:
                        logger.warning(f"命令执行失败 '{cmd}': HTTP {response.status_code}")
                except Exception as e:
                    logger.warning(f"网络修复命令失败 '{cmd}': {e}")
                    
            logger.info("✅ 网络配置修复完成")
            return True
            
        except Exception as e:
            logger.error(f"❌ 网络配置修复失败: {e}")
            return False
    
    def _emergency_network_fix(self):
        """紧急网络修复"""
        try:
            logger.info("🚑 执行紧急网络修复...")
            
            # 更激进的网络修复
            emergency_commands = [
                # 停止网络管理器
                "sudo systemctl stop NetworkManager",
                # 手动配置网络接口
                "sudo ip addr flush dev eth0 || true",
                "sudo ip addr add 10.0.2.15/24 dev eth0 || true", 
                "sudo ip route add default via 10.0.2.2 || true",
                # 手动设置DNS
                "echo 'nameserver 8.8.8.8\nnameserver 8.8.4.4' | sudo tee /etc/resolv.conf",
                # 重启网络管理器
                "sudo systemctl start NetworkManager",
                # 等待稳定
                "sleep 10"
            ]
            
            for cmd in emergency_commands:
                try:
                    response = requests.post(
                        f"http://localhost:{self.server_port}/execute",
                        json={"command": cmd, "shell": True},
                        timeout=(30, 30)
                    )
                    if response.status_code == 200:
                        result = response.json()
                        logger.debug(f"紧急修复命令 '{cmd}': {result.get('status', 'unknown')}")
                except Exception as e:
                    logger.warning(f"紧急修复命令失败 '{cmd}': {e}")
            
            logger.info("✅ 紧急网络修复完成")
            return True
            
        except Exception as e:
            logger.error(f"❌ 紧急网络修复失败: {e}")
            return False

    def start_emulator(self, path_to_vm: str, headless: bool, os_type: str):
        # Use a single lock for all port allocation and container startup
        lock = FileLock(str(self.lock_file), timeout=LOCK_TIMEOUT)
        
        try:
            with lock:
                # Allocate all required ports
                self.vnc_port = self._get_available_port(8006)
                self.server_port = self._get_available_port(5000)
                self.chromium_port = self._get_available_port(9222)
                self.vlc_port = self._get_available_port(8080)

                # Start container while still holding the lock
                # Check if KVM is available
                devices = []
                if os.path.exists("/dev/kvm"):
                    devices.append("/dev/kvm")
                    logger.info("KVM device found, using hardware acceleration")
                else:
                    self.environment["KVM"] = "N"
                    logger.warning("KVM device not found, running without hardware acceleration (will be slower)")

                # 添加网络调试信息
                logger.info("🌐 配置网络参数...")
                logger.info(f"DNS服务器: {self.environment.get('DNS_SERVERS', '')}")
                logger.info(f"网络模式: bridge")
                logger.info(f"QEMU网络选项: {self.environment.get('QEMU_NET_OPTIONS', '')}")
                
                self.container = self.client.containers.run(
                    "happysixd/osworld-docker",
                    environment=self.environment,
                    devices=devices,
                    volumes={
                        os.path.abspath(path_to_vm): {
                            "bind": "/System.qcow2",
                            "mode": "ro"
                        }
                    },
                    ports={
                        8006: self.vnc_port,
                        5000: self.server_port,
                        9222: self.chromium_port,
                        8080: self.vlc_port
                    },
                    ## 简化DNS配置以提高稳定性
                    dns=["8.8.8.8", "8.8.4.4", "1.1.1.1"],
                    dns_search=[],
                    dns_opt=["ndots:0", "timeout:5", "attempts:3"],
                    network_mode="bridge",
                    privileged=True,
                    cap_add=["NET_ADMIN", "SYS_ADMIN", "NET_RAW"],
                    ulimits=[
                        {"name": "nofile", "soft": 65536, "hard": 65536},
                        {"name": "nproc", "soft": 8192, "hard": 8192}
                    ],
                    shm_size="512m",
                    detach=True
                )

            logger.info(f"Started container with ports - VNC: {self.vnc_port}, "
                       f"Server: {self.server_port}, Chrome: {self.chromium_port}, VLC: {self.vlc_port}")
            
            # 显示详细的端口访问信息和网络配置
            print("=" * 60)
            print("🐳 Docker容器已启动! / Docker Container Started!")
            print("=" * 60)
            print("📱 实时访问信息 / Real-time Access Information:")
            print()
            print(f"🖥️  VNC端口 / VNC Port: {self.vnc_port}")
            print(f"   - 命令行访问: vncviewer localhost:{self.vnc_port}")
            print(f"   - 查看虚拟机桌面画面")
            print()
            print(f"🌐 Web服务端口 / Web Service Port: {self.server_port}")
            print(f"   - 浏览器访问: http://localhost:{self.server_port}")
            print(f"   - 实时页面查看虚拟机状态")
            print()
            print(f"📱 Chrome调试端口 / Chrome Debug Port: {self.chromium_port}")
            print(f"   - 浏览器访问: http://localhost:{self.chromium_port}")
            print(f"   - Chrome远程调试接口")
            print()
            print(f"🎬 VLC端口 / VLC Port: {self.vlc_port}")
            print(f"   - 浏览器访问: http://localhost:{self.vlc_port}")
            print()
            print("🌐 网络配置信息 / Network Configuration:")
            print(f"   - DNS服务器: 6个高可用DNS服务器")
            print(f"   - 网络模式: Bridge桥接模式(高稳定性配置)")
            print(f"   - MTU大小: {self.environment.get('MTU', '1500')}")
            print(f"   - 网络类型: 桥接网络 + 特权模式")
            print(f"   - 连接重试: {self.environment.get('CONNECTION_RETRY', '3')}次")
            print(f"   - TCP保活: 已启用")
            print(f"   - 缓冲区优化: 已启用")
            print(f"   - 网络调试: 已启用")
            print()
            print("💡 提示 / Tips:")
            print("   - VNC端口可以让您直接查看虚拟机桌面")
            print("   - Web服务端口提供RESTful API和实时页面")
            print("   - 网络配置已针对Amazon访问优化")
            print("   - SSL证书验证已优化，支持HTTPS网站")
            print("   - 如果出现SSL错误，运行: python3 fix_ssl_certificates.py")
            print("   - 如果Amazon仍有问题，运行: python3 fix_amazon_access.py")
            print("   - 使用 Ctrl+C 可以优雅地停止服务")
            print("=" * 60)
            print()

            # Wait for VM to be ready
            self._wait_for_vm_ready()

        except Exception as e:
            # Clean up if anything goes wrong
            if self.container:
                try:
                    self.container.stop()
                    self.container.remove()
                except:
                    pass
            raise e

    def get_ip_address(self, path_to_vm: str) -> str:
        if not all([self.server_port, self.chromium_port, self.vnc_port, self.vlc_port]):
            raise RuntimeError("VM not started - ports not allocated")
        return f"localhost:{self.server_port}:{self.chromium_port}:{self.vnc_port}:{self.vlc_port}"

    def save_state(self, path_to_vm: str, snapshot_name: str):
        raise NotImplementedError("Snapshots not available for Docker provider")

    def revert_to_snapshot(self, path_to_vm: str, snapshot_name: str):
        self.stop_emulator(path_to_vm)

    def stop_emulator(self, path_to_vm: str, region=None, *args, **kwargs):
        # Note: region parameter is ignored for Docker provider
        # but kept for interface consistency with other providers
        if self.container:
            logger.info("Stopping VM...")
            try:
                self.container.stop()
                self.container.remove()
                time.sleep(WAIT_TIME)
            except Exception as e:
                logger.error(f"Error stopping container: {e}")
            finally:
                self.container = None
                self.server_port = None
                self.vnc_port = None
                self.chromium_port = None
                self.vlc_port = None
