import os
import platform
import zipfile

from time import sleep
import requests
from tqdm import tqdm

import logging

from desktop_env.providers.base import VMManager

logger = logging.getLogger("desktopenv.providers.docker.DockerVMManager")
logger.setLevel(logging.INFO)

MAX_RETRY_TIMES = 10
RETRY_INTERVAL = 5

UBUNTU_X86_URL = "https://huggingface.co/datasets/xlangai/ubuntu_osworld/resolve/main/Ubuntu.qcow2.zip"
WINDOWS_X86_URL = "https://huggingface.co/datasets/xlangai/windows_osworld/resolve/main/Windows-10-x64.qcow2.zip"
VMS_DIR = "./docker_vm_data"

URL = UBUNTU_X86_URL
DOWNLOADED_FILE_NAME = URL.split('/')[-1]

if platform.system() == 'Windows':
    docker_path = r"C:\Program Files\Docker\Docker"
    os.environ["PATH"] += os.pathsep + docker_path


def _download_vm(vms_dir: str):
    global URL, DOWNLOADED_FILE_NAME
    
    # Check for HF_ENDPOINT environment variable and replace domain if set to hf-mirror.com
    hf_endpoint = os.environ.get('HF_ENDPOINT')
    if hf_endpoint and 'hf-mirror.com' in hf_endpoint:
        URL = URL.replace('huggingface.co', 'hf-mirror.com')
        logger.info(f"Using HF mirror: {URL}")

    downloaded_file_name = DOWNLOADED_FILE_NAME
    downloaded_file_path = os.path.join(vms_dir, downloaded_file_name)
    
    # Check if extracted VM already exists
    if downloaded_file_name.endswith(".zip"):
        vm_name = downloaded_file_name[:-4]  # Remove .zip extension
        extracted_vm_path = os.path.join(vms_dir, vm_name)
        
        if os.path.exists(extracted_vm_path):
            logger.info(f"✅ VM image already exists: {extracted_vm_path}")
            logger.info("Skipping download. If you want to re-download, please delete the existing file.")
            return
    
    # Check if zip file already exists and is complete
    if os.path.exists(downloaded_file_path):
        logger.info(f"Found existing zip file: {downloaded_file_path}")
        try:
            # Test if zip file is valid and complete
            with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
                zip_ref.testzip()  # This will raise an exception if zip is corrupted
            logger.info("✅ Existing zip file is valid. Extracting...")
            
            # Extract the existing valid zip file
            if downloaded_file_name.endswith(".zip"):
                logger.info("Extracting the existing zip file...☕️")
                with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
                    zip_ref.extractall(vms_dir)
                logger.info("Files have been successfully extracted to the directory: " + str(vms_dir))
            return
        except (zipfile.BadZipFile, zipfile.LargeZipFile) as e:
            logger.warning(f"Existing zip file is corrupted: {e}")
            logger.info("Removing corrupted file and re-downloading...")
            os.remove(downloaded_file_path)

    # Download the virtual machine image
    logger.info("Downloading the virtual machine image...")
    downloaded_size = 0

    os.makedirs(vms_dir, exist_ok=True)

    while True:
        headers = {}
        if os.path.exists(downloaded_file_path):
            downloaded_size = os.path.getsize(downloaded_file_path)
            headers["Range"] = f"bytes={downloaded_size}-"

        with requests.get(URL, headers=headers, stream=True) as response:
            if response.status_code == 416:
                # This means the range was not satisfiable, possibly the file was fully downloaded
                logger.info("Fully downloaded or the file size changed.")
                break

            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))

            with open(downloaded_file_path, "ab") as file, tqdm(
                    desc="Progress",
                    total=total_size,
                    unit='iB',
                    unit_scale=True,
                    unit_divisor=1024,
                    initial=downloaded_size,
                    ascii=True
            ) as progress_bar:
                try:
                    for data in response.iter_content(chunk_size=1024):
                        size = file.write(data)
                        progress_bar.update(size)
                except (requests.exceptions.RequestException, IOError) as e:
                    logger.error(f"Download error: {e}")
                    sleep(RETRY_INTERVAL)
                    logger.error("Retrying...")
                else:
                    logger.info("Download succeeds.")
                    break  # Download completed successfully

    if downloaded_file_name.endswith(".zip"):
        # Unzip the downloaded file
        logger.info("Unzipping the downloaded file...☕️")
        with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
            zip_ref.extractall(vms_dir)
        logger.info("Files have been successfully extracted to the directory: " + str(vms_dir))


class DockerVMManager(VMManager):
    def __init__(self, registry_path=""):
        pass

    def add_vm(self, vm_path):
        pass

    def check_and_clean(self):
        pass

    def delete_vm(self, vm_path, region=None, **kwargs):
        # Fixed: Added region and **kwargs parameters for interface compatibility
        pass

    def initialize_registry(self):
        pass

    def list_free_vms(self):
        return os.path.join(VMS_DIR, DOWNLOADED_FILE_NAME)

    def occupy_vm(self, vm_path, pid, region=None, **kwargs):
        # Fixed: Added pid, region and **kwargs parameters for interface compatibility
        pass

    def get_vm_path(self, os_type, region, screen_size=(1920, 1080), **kwargs):
        # Note: screen_size parameter is ignored for Docker provider
        # but kept for interface consistency with other providers
        global URL, DOWNLOADED_FILE_NAME
        if os_type == "Ubuntu":
            URL = UBUNTU_X86_URL
        elif os_type == "Windows":
            URL = WINDOWS_X86_URL
        
        # Check for HF_ENDPOINT environment variable and replace domain if set to hf-mirror.com
        hf_endpoint = os.environ.get('HF_ENDPOINT')
        if hf_endpoint and 'hf-mirror.com' in hf_endpoint:
            URL = URL.replace('huggingface.co', 'hf-mirror.com')
            logger.info(f"Using HF mirror: {URL}")
            
        DOWNLOADED_FILE_NAME = URL.split('/')[-1]

        if DOWNLOADED_FILE_NAME.endswith(".zip"):
            vm_name = DOWNLOADED_FILE_NAME[:-4]
        else:
            vm_name = DOWNLOADED_FILE_NAME

        vm_path = os.path.join(VMS_DIR, vm_name)
        zip_path = os.path.join(VMS_DIR, DOWNLOADED_FILE_NAME)
        
        # Check if VM file exists (either extracted .qcow2 or zip file)
        if not os.path.exists(vm_path):
            if os.path.exists(zip_path):
                logger.info(f"Found existing zip file: {zip_path}")
                logger.info("Extracting to get VM file...")
            _download_vm(VMS_DIR)
        else:
            logger.info(f"✅ Using existing VM image: {vm_path}")
            
        return vm_path
