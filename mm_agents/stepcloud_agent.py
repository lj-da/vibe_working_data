import base64
import json
import logging
import time
import os
from io import BytesIO
from typing import Dict, List, Tuple
from PIL import Image

# 导入 stepcloud_api_qwen 中的 step_func
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from stepcloud_api_qwen import step_func
except ImportError:
    # 如果直接导入失败，尝试相对路径导入
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from stepcloud_api_qwen import step_func

logger = None

def process_image(image_bytes):
    """
    Process an image for the model.
    Resize the image to dimensions expected by the model.
    
    Args:
        image_bytes: Raw image bytes
    
    Returns:
        Base64 encoded image string of the processed image
    """
    # Open image from bytes
    image = Image.open(BytesIO(image_bytes))
    width, height = image.size
    
    # 保持原始尺寸，不进行缩放
    # Convert to bytes
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    processed_bytes = buffer.getvalue()
    
    # Return base64 encoded string
    return base64.b64encode(processed_bytes).decode('utf-8')


class StepCloudAgent:
    """
    Agent that uses StepCloud API for reasoning and action generation
    """

    def __init__(
        self,
        model_name="ovr_zero_dogegg_g5_fixres_mount_rftdata_it500_out",
        max_tokens=8192,
        top_p=0.9,
        temperature=0.2,
        action_space="pyautogui",
        observation_type="screenshot",
        history_n=4,
        add_thought_prefix=False,
    ):
        self.model_name = model_name
        self.max_tokens = max_tokens
        self.top_p = top_p
        self.temperature = temperature
        self.action_space = action_space
        self.observation_type = observation_type
        self.history_n = history_n
        self.add_thought_prefix = add_thought_prefix
        assert action_space in ["pyautogui"], "Invalid action space"
        assert observation_type in ["screenshot"], "Invalid observation type"
        
        self.thoughts = []
        self.actions = []
        self.observations = []
        self.responses = []
        self.screenshots = []

    def predict(self, instruction: str, obs: Dict) -> Tuple[str, List[str]]:
        """
        Predict the next action(s) based on the current observation using StepCloud API.
        
        Args:
            instruction: The task instruction
            obs: Current observation containing screenshot
            
        Returns:
            Tuple of (response_text, list_of_actions)
        """
        # Process the screenshot image
        screenshot_bytes = obs["screenshot"]
        
        # Display original dimensions
        image = Image.open(BytesIO(screenshot_bytes))
        width, height = image.size
        print(f"Original screen resolution: {width}x{height}")
        
        # Process the image
        processed_image = process_image(screenshot_bytes)
        
        # Save the current screenshot to history
        self.screenshots.append(processed_image)
        
        # Calculate history window start index
        current_step = len(self.actions)
        history_start_idx = max(0, current_step - self.history_n)
        
        # Build previous actions string - only include actions outside the history window
        previous_actions = []
        for i in range(history_start_idx):
            if i < len(self.actions):
                previous_actions.append(f"Step {i+1}: {self.actions[i]}")
        previous_actions_str = "\n".join(previous_actions) if previous_actions else "None"

        # Create instruction prompt
        instruction_prompt = f"""
Please generate the next move according to the UI screenshot, instruction and previous actions.

Instruction: {instruction}

Previous actions:
{previous_actions_str}

Please provide your response in the following format:
1. First, analyze the current screenshot and understand the context
2. Then, provide the next action to take
3. Finally, output the pyautogui code to execute the action

The action should be one of:
- pyautogui.click(x, y) - for left click at coordinates
- pyautogui.rightClick(x, y) - for right click at coordinates
- pyautogui.doubleClick(x, y) - for double click at coordinates
- pyautogui.typewrite(text) - for typing text
- pyautogui.press(key) - for pressing a key
- pyautogui.hotkey(key1, key2) - for key combinations
- time.sleep(seconds) - for waiting
- pyautogui.scroll(pixels) - for scrolling

Please provide the coordinates based on the screenshot dimensions: {width}x{height}
"""

        # Prepare images for step_func
        images = [processed_image]
        
        # Call StepCloud API using step_func
        try:
            response = step_func(
                model_name=self.model_name,
                images=images,
                prompt=instruction_prompt
            )
        except Exception as e:
            logger.error(f"Error calling StepCloud API: {e}")
            response = f"Error: {str(e)}"
        
        logger.info(f"StepCloud Output: {response}")
        
        # Save response to history
        self.responses.append(response)
        
        # Parse response and extract pyautogui code
        low_level_instruction, pyautogui_code = self.parse_response(
            response, 
            width, 
            height
        )

        logger.info(f"Low level instruction: {low_level_instruction}")
        logger.info(f"Pyautogui code: {pyautogui_code}")

        # Add the action to history
        self.actions.append(low_level_instruction)

        return response, pyautogui_code

    def parse_response(self, response: str, original_width: int = None, original_height: int = None) -> Tuple[str, List[str]]:
        """
        Parse StepCloud response and convert it to low level action and pyautogui code.
        
        Args:
            response: Raw response string from the model
            original_width: Width of the original screenshot
            original_height: Height of the original screenshot
            
        Returns:
            Tuple of (low_level_instruction, list of pyautogui_commands)
        """
        low_level_instruction = ""
        pyautogui_code = []
        
        if response is None or not response.strip():
            return low_level_instruction, pyautogui_code
        
        # 简单的解析逻辑，提取 pyautogui 代码
        lines = response.split('\n')
        for line in lines:
            line = line.strip()
            if line.startswith('pyautogui.') or line.startswith('time.sleep('):
                pyautogui_code.append(line)
            elif 'click' in line.lower() or 'type' in line.lower() or 'press' in line.lower():
                # 提取动作描述
                if not low_level_instruction:
                    low_level_instruction = line
        
        # 如果没有找到 pyautogui 代码，尝试从响应中生成
        if not pyautogui_code and response:
            # 简单的启发式方法：如果响应包含坐标信息，生成点击动作
            import re
            coord_match = re.search(r'(\d+)\s*[,\s]\s*(\d+)', response)
            if coord_match:
                x, y = int(coord_match.group(1)), int(coord_match.group(2))
                pyautogui_code.append(f"pyautogui.click({x}, {y})")
                low_level_instruction = f"Click at coordinates ({x}, {y})"
        
        # 如果仍然没有动作，生成默认动作
        if not low_level_instruction and len(pyautogui_code) > 0:
            action_type = pyautogui_code[0].split(".", 1)[1].split("(", 1)[0]
            low_level_instruction = f"Performing {action_type} action"
        
        return low_level_instruction, pyautogui_code

    def reset(self, _logger=None):
        global logger
        logger = (_logger if _logger is not None else
                  logging.getLogger("desktopenv.stepcloud_agent"))

        self.thoughts = []
        self.action_descriptions = []
        self.actions = []
        self.observations = []
        self.responses = []
        self.screenshots = []
