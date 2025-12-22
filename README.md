
## 💾 Installation
### VMware/VirtualBox (Desktop, Laptop, Bare Metal Machine)
Suppose you are operating on a system that has not been virtualized (e.g. your desktop, laptop, bare metal machine), meaning you are not utilizing a virtualized environment like AWS, Azure, or k8s.
If this is the case, proceed with the instructions below. However, if you are on a virtualized platform, please refer to the [Docker](https://github.com/xlang-ai/OSWorld?tab=readme-ov-file#docker-server-with-kvm-support-for-the-better) section.

1. First, clone this repository and `cd` into it. Then, install the dependencies listed in `requirements.txt`. It is recommended that you use the latest version of Conda to manage the environment, but you can also choose to manually install the dependencies. Please ensure that the version of Python is >= 3.10.
```bash
# Clone the OSWorld repository
git clone https://github.com/xlang-ai/OSWorld

# Change directory into the cloned repository
cd OSWorld

# Optional: Create a Conda environment for OSWorld
# conda create -n osworld python=3.10
# conda activate osworld

# Install required dependencies
pip install -r requirements.txt
```

Alternatively, you can install the environment without any benchmark tasks:
```bash
pip install desktop-env
```

2. Install [VMware Workstation Pro](https://www.vmware.com/products/workstation-pro/workstation-pro-evaluation.html) (for systems with Apple Chips, you should install [VMware Fusion](https://support.broadcom.com/group/ecx/productdownloads?subfamily=VMware+Fusion)) and configure the `vmrun` command.  The installation process can refer to [How to install VMware Workstation Pro](desktop_env/providers/vmware/INSTALL_VMWARE.md). Verify the successful installation by running the following:
```bash
vmrun -T ws list
```
If the installation along with the environment variable set is successful, you will see the message showing the current running virtual machines.
> **Note:** We also support using [VirtualBox](https://www.virtualbox.org/) if you have issues with VMware Pro. However, features such as parallelism and macOS on Apple chips might not be well-supported.

All set! Our setup script will automatically download the necessary virtual machines and configure the environment for you.

### Docker (Server with KVM Support for Better Performance)
If you are running on a non-bare metal server, or prefer not to use VMware and VirtualBox platforms, we recommend using our Docker support.

#### Prerequisite: Check if your machine supports KVM
We recommend running the VM with KVM support. To check if your hosting platform supports KVM, run
```
egrep -c '(vmx|svm)' /proc/cpuinfo
```
on Linux. If the return value is greater than zero, the processor should be able to support KVM.
> **Note**: macOS hosts generally do not support KVM. You are advised to use VMware if you would like to run OSWorld on macOS.

#### Install Docker
If your hosting platform supports a graphical user interface (GUI), you may refer to [Install Docker Desktop on Linux](https://docs.docker.com/desktop/install/linux/) or [Install Docker Desktop on Windows](https://docs.docker.com/desktop/install/windows-install/) based on your OS. Otherwise, you may [Install Docker Engine](https://docs.docker.com/engine/install/).

#### Running Experiments
Add the following arguments when initializing `DesktopEnv`: 
- `provider_name`: `docker`
- `os_type`: `Ubuntu` or `Windows`, depending on the OS of the VM
> **Note**: If the experiment is interrupted abnormally (e.g., by interrupting signals), there may be residual docker containers which could affect system performance over time. Please run `docker stop $(docker ps -q) && docker rm $(docker ps -a -q)` to clean up.
> Docker VM Data已经SYNC到s3://shenyeqing/project/GUI/osworld/docker_vm_data/ 进行备份


#### Docker环境测试 / Docker Environment Testing
使用以下命令测试Docker环境是否能正常启动：

```bash
# 启动Docker虚拟机并启用GUI
python3 run_docker.py --enable_gui --task_config dock_example/example_task_config.json
```

启动成功后，程序会在日志中显示VNC端口信息，类似于：
```
🖥️  VNC端口 / VNC Port: 8008
```

#### 通过VNC查看虚拟机页面 / Accessing VM via VNC
找到日志中的VNC端口后，可以通过以下方式查看虚拟机桌面：
**使用Web浏览器**
某些配置下也可以通过浏览器访问：
```
http://localhost:8008
```

> **注意**: VNC端口通常在8006-8010之间，具体端口号以程序日志中显示的为准。

#### 故障排除 / Troubleshooting
如果Docker无法正常启动，请检查：
1. Docker服务是否正在运行：`sudo systemctl status docker`
2. 当前用户是否在docker组中：`groups $USER`
3. 系统是否支持KVM：`egrep -c '(vmx|svm)' /proc/cpuinfo`
4. 清理残留容器：`docker stop $(docker ps -q) && docker rm $(docker ps -a -q)`

#### Step模型OSWorld测试 / Step Model OSWorld Testing
使用以下命令启动Step模型的OSWorld测试，支持并行环境执行：

```bash
# 启动Step模型OSWorld测试
python3 run_multienv_stepcopilot_v1.py \
    --provider_name docker \
    --enable_gui \
    --model [stepcloud上部署的模型名称] \
    --num_envs [并行数量]
```

**参数说明 / Parameter Description:**
- `--provider_name docker`: 使用Docker环境
- `--enable_gui`: 启用GUI显示，便于通过VNC查看执行过程
- `--model`: stepcloud上部署的模型名称，例如：`cu_sft_0902_nothink_history_508`
- `--num_envs`: 并行执行的环境数量，例如：`2` 或 `4`

**使用示例 / Usage Examples:**

```bash
# 示例1: 使用2个并行环境测试
python3 run_multienv_stepcopilot_v1.py \
    --provider_name docker \
    --enable_gui \
    --model cu_sft_0902_nothink_history_508 \
    --num_envs 2

# 示例2: 使用4个并行环境，指定测试域
python3 run_multienv_stepcopilot_v1.py \
    --provider_name docker \
    --enable_gui \
    --model your_stepcloud_model_name \
    --num_envs 4 \
    --domain chrome
```

**高级配置选项 / Advanced Configuration Options:**
- `--domain`: 指定测试域（例如：`chrome`, `libreoffice_calc`, `all`）
- `--max_steps`: 最大执行步数（默认：30）
- `--temperature`: 模型采样温度（默认：1.0）
- `--max_tokens`: 最大生成token数（默认：500）
- `--result_dir`: 结果保存目录（默认：`./results`）

**监控并行执行 / Monitoring Parallel Execution:**
启动后，每个并行环境都会分配独立的VNC端口（通常从8006开始递增）：
```
🖥️  Environment 1 VNC Port: 8006
🖥️  Environment 2 VNC Port: 8007
🖥️  Environment 3 VNC Port: 8008
🖥️  Environment 4 VNC Port: 8009
```

通过以下命令分别查看各个环境的执行状态：
```bash
# 查看环境1
vncviewer localhost:8006

# 查看环境2  
vncviewer localhost:8007
```

> **注意**: 并行环境数量建议根据系统资源调整。每个环境大约需要2-4GB内存和一定的CPU资源。

#### 计算测试分数 / Calculate Test Scores
测试完成后，使用以下命令计算模型的分数：

```bash
python3 simple_score.py --model_name [模型名称]
```

**参数说明 / Parameter Description:**
- `--model_name`: 指定要计算分数的模型名称，例如：`cu_sft_0902_nothink_history_508`

**使用示例 / Usage Example:**
```bash
# 计算指定模型的测试分数
python3 simple_score.py --model_name cu_sft_0902_nothink_history_508
```


