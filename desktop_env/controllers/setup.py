import json
import logging
import os
import os.path
import platform
import shutil
import sqlite3
import tempfile
import time
import traceback
import uuid
from datetime import datetime, timedelta
from typing import Any, Union, Optional
from typing import Dict, List

import requests
from playwright.sync_api import sync_playwright, TimeoutError
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive, GoogleDriveFile, GoogleDriveFileList
from requests_toolbelt.multipart.encoder import MultipartEncoder

from desktop_env.controllers.python import PythonController
from desktop_env.evaluators.metrics.utils import compare_urls
from desktop_env.providers.aws.proxy_pool import get_global_proxy_pool, init_proxy_pool, ProxyInfo

import dotenv
# Load environment variables from .env file
dotenv.load_dotenv()


PROXY_CONFIG_FILE = os.getenv("PROXY_CONFIG_FILE", "evaluation_examples/settings/proxy/dataimpulse.json")  # Default proxy config file

logger = logging.getLogger("desktopenv.setup")

FILE_PATH = os.path.dirname(os.path.abspath(__file__))

# init_proxy_pool(PROXY_CONFIG_FILE)  # initialize the global proxy pool - DISABLED for stability

MAX_RETRIES = 20

class SetupController:
    def __init__(self, vm_ip: str, server_port: int = 5000, chromium_port: int = 9222, vlc_port: int = 8080, cache_dir: str = "cache", client_password: str = "", screen_width: int = 1920, screen_height: int = 1080):
        self.vm_ip: str = vm_ip
        self.server_port: int = server_port
        self.chromium_port: int = chromium_port
        self.vlc_port: int = vlc_port
        self.http_server: str = f"http://{vm_ip}:{server_port}"
        self.http_server_setup_root: str = f"http://{vm_ip}:{server_port}/setup"
        self.cache_dir: str = cache_dir
        self.use_proxy: bool = False
        self.client_password: str = client_password
        self.screen_width: int = screen_width
        self.screen_height: int = screen_height

    def reset_cache_dir(self, cache_dir: str):
        self.cache_dir = cache_dir

    def setup(self, config: List[Dict[str, Any]], use_proxy: bool = False)-> bool:
        """
        Args:
            config (List[Dict[str, Any]]): list of dict like {str: Any}. each
              config dict has the structure like
                {
                    "type": str, corresponding to the `_{:}_setup` methods of
                      this class
                    "parameters": dict like {str, Any} providing the keyword
                      parameters
                }
        """  
        self.use_proxy = use_proxy
        # make sure connection can be established
        logger.info(f"try to connect {self.http_server}")
        retry = 0
        while retry < MAX_RETRIES:
            try:
                _ = requests.get(self.http_server + "/terminal")
                break
            except:
                time.sleep(5)
                retry += 1
                logger.info(f"retry: {retry}/{MAX_RETRIES}")
            
            if retry == MAX_RETRIES:
                return False
                

        for i, cfg in enumerate(config):
            config_type: str = cfg["type"]
            parameters: Dict[str, Any] = cfg["parameters"]

            # Assumes all the setup the functions should follow this name
            # protocol
            setup_function: str = "_{:}_setup".format(config_type)
            assert hasattr(self, setup_function), f'Setup controller cannot find init function {setup_function}'
            
            try:
                logger.info(f"Executing setup step {i+1}/{len(config)}: {setup_function}")
                logger.debug(f"Setup parameters: {parameters}")
                getattr(self, setup_function)(**parameters)
                logger.info(f"SETUP COMPLETED: {setup_function}({str(parameters)})")
            except Exception as e:
                logger.error(f"SETUP FAILED at step {i+1}/{len(config)}: {setup_function}({str(parameters)})")
                logger.error(f"Error details: {e}")
                logger.error(f"Traceback: {traceback.format_exc()}")
                raise Exception(f"Setup step {i+1} failed: {setup_function} - {e}") from e
        
        return True

    def _download_setup(self, files: List[Dict[str, str]]):
        """
        Args:
            files (List[Dict[str, str]]): files to download. lisf of dict like
              {
                "url": str, the url to download
                "path": str, the path on the VM to store the downloaded file
              }
        """
        for f in files:
            url: str = f["url"]
            path: str = f["path"]
            cache_path: str = os.path.join(self.cache_dir, "{:}_{:}".format(
                uuid.uuid5(uuid.NAMESPACE_URL, url),
                os.path.basename(path)))
            if not url or not path:
                raise Exception(f"Setup Download - Invalid URL ({url}) or path ({path}).")

            if not os.path.exists(cache_path):
                logger.info(f"Cache file not found, downloading from {url} to {cache_path}")
                max_retries = 3
                downloaded = False
                e = None
                for i in range(max_retries):
                    try:
                        logger.info(f"Download attempt {i+1}/{max_retries} for {url}")
                        response = requests.get(url, stream=True, timeout=300)  # Add 5 minute timeout
                        response.raise_for_status()
                        
                        # Get file size if available
                        total_size = int(response.headers.get('content-length', 0))
                        if total_size > 0:
                            logger.info(f"File size: {total_size / (1024*1024):.2f} MB")

                        downloaded_size = 0
                        with open(cache_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                                    downloaded_size += len(chunk)
                                    if total_size > 0 and downloaded_size % (1024*1024) == 0:  # Log every MB
                                        progress = (downloaded_size / total_size) * 100
                                        logger.info(f"Download progress: {progress:.1f}%")
                        
                        logger.info(f"File downloaded successfully to {cache_path} ({downloaded_size / (1024*1024):.2f} MB)")
                        downloaded = True
                        break

                    except requests.RequestException as e:
                        logger.error(
                            f"Failed to download {url} caused by {e}. Retrying... ({max_retries - i - 1} attempts left)")
                        # Clean up partial download
                        if os.path.exists(cache_path):
                            os.remove(cache_path)
                if not downloaded:
                    raise requests.RequestException(f"Failed to download {url}. No retries left.")

            form = MultipartEncoder({
                "file_path": path,
                "file_data": (os.path.basename(path), open(cache_path, "rb"))
            })
            headers = {"Content-Type": form.content_type}
            logger.debug(form.content_type)

            # send request to server to upload file
            try:
                logger.info(f"Uploading {os.path.basename(path)} to VM at {path}")
                logger.debug("REQUEST ADDRESS: %s", self.http_server + "/setup" + "/upload")
                response = requests.post(self.http_server + "/setup" + "/upload", headers=headers, data=form, timeout=600)  # 10 minute timeout for upload
                if response.status_code == 200:
                    logger.info(f"File uploaded successfully: {path}")
                    logger.debug("Upload response: %s", response.text)
                else:
                    logger.error(f"Failed to upload file {path}. Status code: {response.status_code}, Response: {response.text}")
                    raise requests.RequestException(f"Upload failed with status {response.status_code}")
            except requests.exceptions.RequestException as e:
                logger.error(f"An error occurred while trying to upload {path}: {e}")
                raise

    def _upload_file_setup(self, files: List[Dict[str, str]]):
        """
        Args:
            files (List[Dict[str, str]]): files to download. lisf of dict like
              {
                "local_path": str, the local path to the file to upload
                "path": str, the path on the VM to store the downloaded file
              }
        """
        for f in files:
            local_path: str = f["local_path"]
            path: str = f["path"]

            if not os.path.exists(local_path):
                logger.error(f"Setup Upload - Invalid local path ({local_path}).")
                return

            form = MultipartEncoder({
                "file_path": path,
                "file_data": (os.path.basename(path), open(local_path, "rb"))
            })
            headers = {"Content-Type": form.content_type}
            logger.debug(form.content_type)

            # send request to server to upload file
            try:
                logger.debug("REQUEST ADDRESS: %s", self.http_server + "/setup" + "/upload")
                response = requests.post(self.http_server + "/setup" + "/upload", headers=headers, data=form)
                if response.status_code == 200:
                    logger.info("Command executed successfully: %s", response.text)
                else:
                    logger.error("Failed to upload file. Status code: %s", response.text)
            except requests.exceptions.RequestException as e:
                logger.error("An error occurred while trying to send the request: %s", e)

    def _change_wallpaper_setup(self, path: str):
        if not path:
            raise Exception(f"Setup Wallpaper - Invalid path ({path}).")

        payload = json.dumps({"path": path})
        headers = {
            'Content-Type': 'application/json'
        }

        # send request to server to change wallpaper
        try:
            response = requests.post(self.http_server + "/setup" + "/change_wallpaper", headers=headers, data=payload)
            if response.status_code == 200:
                logger.info("Command executed successfully: %s", response.text)
            else:
                logger.error("Failed to change wallpaper. Status code: %s", response.text)
        except requests.exceptions.RequestException as e:
            logger.error("An error occurred while trying to send the request: %s", e)

    def _tidy_desktop_setup(self, **config):
        raise NotImplementedError()

    def _open_setup(self, path: str):
        if not path:
            raise Exception(f"Setup Open - Invalid path ({path}).")

        payload = json.dumps({"path": path})
        headers = {
            'Content-Type': 'application/json'
        }

        # send request to server to open file
        try:
            # The server-side call is now blocking and can take time.
            # We set a timeout that is slightly longer than the server's timeout (1800s).
            response = requests.post(self.http_server + "/setup" + "/open_file", headers=headers, data=payload, timeout=1810)
            response.raise_for_status()  # This will raise an exception for 4xx and 5xx status codes
            logger.info("Command executed successfully: %s", response.text)
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to open file '{path}'. An error occurred while trying to send the request or the server responded with an error: {e}")
            raise Exception(f"Failed to open file '{path}'. An error occurred while trying to send the request or the server responded with an error: {e}") from e

    def _launch_setup(self, command: Union[str, List[str]], shell: bool = False):
        if not command:
            raise Exception("Empty command to launch.")

        if not shell and isinstance(command, str) and len(command.split()) > 1:
            logger.warning("Command should be a list of strings. Now it is a string. Will split it by space.")
            command = command.split()
            
        # 代理配置已禁用以提高网络稳定性
        # if command[0] == "google-chrome" and self.use_proxy:
        #     command.append("--proxy-server=http://127.0.0.1:18888")  # Use the proxy server set up by _proxy_setup

        payload = json.dumps({"command": command, "shell": shell})
        headers = {"Content-Type": "application/json"}

        try:
            logger.info("REQUEST ADDRESS: %s", self.http_server + "/setup" + "/launch")
            response = requests.post(self.http_server + "/setup" + "/launch", headers=headers, data=payload)
            if response.status_code == 200:
                logger.info("Command executed successfully: %s", response.text)
            else:
                logger.error("Failed to launch application. Status code: %s", response.text)
        except requests.exceptions.RequestException as e:
            logger.error("An error occurred while trying to send the request: %s", e)

    def _execute_setup(
            self,
            command: List[str],
            stdout: str = "",
            stderr: str = "",
            shell: bool = False,
            until: Optional[Dict[str, Any]] = None
    ):
        if not command:
            raise Exception("Empty command to launch.")

        until: Dict[str, Any] = until or {}
        terminates: bool = False
        nb_failings = 0

        def replace_screen_env_in_command(command):
            password = self.client_password
            width = self.screen_width
            height = self.screen_height
            width_half = str(width // 2)
            height_half = str(height // 2)
            new_command_list = []
            new_command = ""
            if isinstance(command, str):
                new_command = command.replace("{CLIENT_PASSWORD}", password)
                new_command = new_command.replace("{SCREEN_WIDTH_HALF}", width_half)
                new_command = new_command.replace("{SCREEN_HEIGHT_HALF}", height_half)
                new_command = new_command.replace("{SCREEN_WIDTH}", str(width))
                new_command = new_command.replace("{SCREEN_HEIGHT}", str(height))
                return new_command
            else:
                for item in command:
                    item = item.replace("{CLIENT_PASSWORD}", password)
                    item = item.replace("{SCREEN_WIDTH_HALF}", width_half)
                    item = item.replace("{SCREEN_HEIGHT_HALF}", height_half)
                    item = item.replace("{SCREEN_WIDTH}", str(width))
                    item = item.replace("{SCREEN_HEIGHT}", str(height))
                    new_command_list.append(item)
                return new_command_list
        command = replace_screen_env_in_command(command)
        payload = json.dumps({"command": command, "shell": shell})
        headers = {"Content-Type": "application/json"}

        while not terminates:
            try:
                response = requests.post(self.http_server + "/setup" + "/execute", headers=headers, data=payload)
                if response.status_code == 200:
                    results: Dict[str, str] = response.json()
                    if stdout:
                        with open(os.path.join(self.cache_dir, stdout), "w") as f:
                            f.write(results["output"])
                    if stderr:
                        with open(os.path.join(self.cache_dir, stderr), "w") as f:
                            f.write(results["error"])
                    logger.info("Command executed successfully: %s -> %s"
                                , " ".join(command) if isinstance(command, list) else command
                                , response.text
                                )
                else:
                    logger.error("Failed to launch application. Status code: %s", response.text)
                    results = None
                    nb_failings += 1
            except requests.exceptions.RequestException as e:
                logger.error("An error occurred while trying to send the request: %s", e)
                traceback.print_exc()

                results = None
                nb_failings += 1

            if len(until) == 0:
                terminates = True
            elif results is not None:
                terminates = "returncode" in until and results["returncode"] == until["returncode"] \
                             or "stdout" in until and until["stdout"] in results["output"] \
                             or "stderr" in until and until["stderr"] in results["error"]
            terminates = terminates or nb_failings >= 5
            if not terminates:
                time.sleep(0.3)

    def _execute_with_verification_setup(
            self,
            command: List[str],
            verification: Dict[str, Any] = None,
            max_wait_time: int = 10,
            check_interval: float = 1.0,
            shell: bool = False
    ):
        """Execute command with verification of results
        
        Args:
            command: Command to execute
            verification: Dict with verification criteria:
                - window_exists: Check if window with this name exists
                - command_success: Execute this command and check if it succeeds
            max_wait_time: Maximum time to wait for verification
            check_interval: Time between verification checks
            shell: Whether to use shell
        """
        if not command:
            raise Exception("Empty command to launch.")

        verification = verification or {}
        
        payload = json.dumps({
            "command": command, 
            "shell": shell,
            "verification": verification,
            "max_wait_time": max_wait_time,
            "check_interval": check_interval
        })
        headers = {"Content-Type": "application/json"}

        try:
            response = requests.post(self.http_server + "/setup" + "/execute_with_verification", 
                                   headers=headers, data=payload, timeout=max_wait_time + 10)
            if response.status_code == 200:
                result = response.json()
                logger.info("Command executed and verified successfully: %s -> %s"
                            , " ".join(command) if isinstance(command, list) else command
                            , response.text
                            )
                return result
            else:
                logger.error("Failed to execute with verification. Status code: %s", response.text)
                raise Exception(f"Command verification failed: {response.text}")
        except requests.exceptions.RequestException as e:
            logger.error("An error occurred while trying to send the request: %s", e)
            traceback.print_exc()
            raise Exception(f"Request failed: {e}")

    def _command_setup(self, command: List[str], **kwargs):
        self._execute_setup(command, **kwargs)

    def _sleep_setup(self, seconds: float):
        time.sleep(seconds)

    def _act_setup(self, action_seq: List[Union[Dict[str, Any], str]]):
        # TODO
        raise NotImplementedError()

    def _replay_setup(self, trajectory: str):
        """
        Args:
            trajectory (str): path to the replay trajectory file
        """

        # TODO
        raise NotImplementedError()

    def _activate_window_setup(self, window_name: str, strict: bool = False, by_class: bool = False):
        if not window_name:
            raise Exception(f"Setup Open - Invalid path ({window_name}).")

        payload = json.dumps({"window_name": window_name, "strict": strict, "by_class": by_class})
        headers = {
            'Content-Type': 'application/json'
        }

        # send request to server to open file
        try:
            response = requests.post(self.http_server + "/setup" + "/activate_window", headers=headers, data=payload)
            if response.status_code == 200:
                logger.info("Command executed successfully: %s", response.text)
            else:
                logger.error(f"Failed to activate window {window_name}. Status code: %s", response.text)
        except requests.exceptions.RequestException as e:
            logger.error("An error occurred while trying to send the request: %s", e)

    def _close_window_setup(self, window_name: str, strict: bool = False, by_class: bool = False):
        if not window_name:
            raise Exception(f"Setup Open - Invalid path ({window_name}).")

        payload = json.dumps({"window_name": window_name, "strict": strict, "by_class": by_class})
        headers = {
            'Content-Type': 'application/json'
        }

        # send request to server to open file
        try:
            response = requests.post(self.http_server + "/setup" + "/close_window", headers=headers, data=payload)
            if response.status_code == 200:
                logger.info("Command executed successfully: %s", response.text)
            else:
                logger.error(f"Failed to close window {window_name}. Status code: %s", response.text)
        except requests.exceptions.RequestException as e:
            logger.error("An error occurred while trying to send the request: %s", e)

    def _proxy_setup(self, client_password: str = ""):
        """Setup system-wide proxy configuration using proxy pool - DISABLED for network stability
        
        Args:
            client_password (str): Password for sudo operations, defaults to "password"
        """
        logger.info("代理配置已禁用，跳过代理设置以提高网络稳定性")
        
        # 仍然需要检查连接性
        retry = 0
        while retry < MAX_RETRIES:
            try:
                _ = requests.get(self.http_server + "/terminal")
                break
            except:
                time.sleep(5)
                retry += 1
                logger.info(f"retry: {retry}/{MAX_RETRIES}")
            
            if retry == MAX_RETRIES:
                return False
        
        return True

    # Chrome setup
    def _chrome_open_tabs_setup(self, urls_to_open: List[str]):
        host = self.vm_ip
        port = self.chromium_port  # fixme: this port is hard-coded, need to be changed from config file

        remote_debugging_url = f"http://{host}:{port}"
        logger.info("Connect to Chrome @: %s", remote_debugging_url)
        logger.debug("PLAYWRIGHT ENV: %s", repr(os.environ))
        for attempt in range(15):
            if attempt > 0:
                time.sleep(5)

            browser = None
            with sync_playwright() as p:
                try:
                    browser = p.chromium.connect_over_cdp(remote_debugging_url)
                    # break
                except Exception as e:
                    if attempt < 14:
                        logger.error(f"Attempt {attempt + 1}: Failed to connect, retrying. Error: {e}")
                        # time.sleep(10)
                        continue
                    else:
                        logger.error(f"Failed to connect after multiple attempts: {e}")
                        raise e

                if not browser:
                    return

                logger.info("Opening %s...", urls_to_open)
                for i, url in enumerate(urls_to_open):
                    # Use the first context (which should be the only one if using default profile)
                    if i == 0:
                        context = browser.contexts[0]

                    page = context.new_page()  # Create a new page (tab) within the existing context
                    try:
                        page.goto(url, timeout=60000)
                    except:
                        logger.warning("Opening %s exceeds time limit", url)  # only for human test
                    logger.info(f"Opened tab {i + 1}: {url}")

                    if i == 0:
                        # clear the default tab
                        default_page = context.pages[0]
                        default_page.close()

                # Do not close the context or browser; they will remain open after script ends
                return browser, context

    def _chrome_close_tabs_setup(self, urls_to_close: List[str]):
        time.sleep(5)  # Wait for Chrome to finish launching

        host = self.vm_ip
        port = self.chromium_port  # fixme: this port is hard-coded, need to be changed from config file

        remote_debugging_url = f"http://{host}:{port}"
        with sync_playwright() as p:
            browser = None
            for attempt in range(15):
                try:
                    browser = p.chromium.connect_over_cdp(remote_debugging_url)
                    break
                except Exception as e:
                    if attempt < 14:
                        logger.error(f"Attempt {attempt + 1}: Failed to connect, retrying. Error: {e}")
                        time.sleep(5)
                    else:
                        logger.error(f"Failed to connect after multiple attempts: {e}")
                        raise e

            if not browser:
                return

            for i, url in enumerate(urls_to_close):
                # Use the first context (which should be the only one if using default profile)
                if i == 0:
                    context = browser.contexts[0]

                for page in context.pages:

                    # if two urls are the same, close the tab
                    if compare_urls(page.url, url):
                        context.pages.pop(context.pages.index(page))
                        page.close()
                        logger.info(f"Closed tab {i + 1}: {url}")
                        break

            # Do not close the context or browser; they will remain open after script ends
            return browser, context

    # google drive setup
    def _googledrive_setup(self, **config):
        """Google Drive setup - DISABLED to avoid authentication errors"""
        print("Warning: Google Drive setup is disabled to avoid authentication errors")
        return True  # Return success to continue execution