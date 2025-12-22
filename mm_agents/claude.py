import ast
import base64
import math
import re
import xml.etree.ElementTree as ET
from io import BytesIO
from typing import Dict, List
import numpy as np
import base64
from loguru import logger
import os
import re
from io import BytesIO
from typing import Dict, List
from PIL import Image
import requests
import json
import time
from tools.ask_llm import ask_llm_anything
from mm_agents.accessibility_tree_wrap.heuristic_retrieve import (
    filter_nodes,
)

# Accessibility tree namespaces
attributes_ns_ubuntu = "https://accessibility.windows.example.org/ns/attributes"
attributes_ns_windows = "https://accessibility.windows.example.org/ns/attributes"
state_ns_ubuntu = "https://accessibility.ubuntu.example.org/ns/state"
state_ns_windows = "https://accessibility.windows.example.org/ns/state"
component_ns_ubuntu = "https://accessibility.ubuntu.example.org/ns/component"
component_ns_windows = "https://accessibility.windows.example.org/ns/component"
value_ns_ubuntu = "https://accessibility.ubuntu.example.org/ns/value"
value_ns_windows = "https://accessibility.windows.example.org/ns/value"
class_ns_windows = "https://accessibility.windows.example.org/ns/class"
# NOTE: this function is used to generate response from the modelproxy
# def generate_response_v2(messages, model, temperature=0.1, max_tokens=4096):
#     api_key = "ak-cc05af1b0e5f8c96424fb8f6db0dda97"
#     headers = {
#         "Content-Type": "application/json",
#         "Authorization": f"Bearer {api_key}"
#     }
#     payload = {
#         "model": model,
#         "messages": messages,
#         "max_tokens": max_tokens,
#         "temperature": temperature
#     }

#     for num in range(5):
#         try:     
#             response = requests.post("https://models-proxy.stepfun-inc.com/v1/chat/completions", headers=headers, json=payload)
#             if response.status_code == 200:
#                 return (response.json())['choices'][0]['message']['content']
#             else :
#                 if num == 4:
#                     return None
#         except Exception as e:
#             print(f"Error occurred: {e}")
#             return None

STEPCOPILOT_ACTION_SPACE = """
explain:xxx\taction:CLICK\tpoint:x,y\n
explain:xxx\taction:TYPE\tvalue:xxxx\tpoint:x,y\n
explain:xxx\taction:COMPLETE\tvalue:任务完成\n
explain:xxx\taction:WAIT\tvalue:n\n
explain:xxx\taction:SCROLL\tvalue:n\tpoint:x,y\n
explain:xxx\taction:HSCROLL\tvalue:n\tpoint:x,y\n
explain:xxx\taction:HOTKEY\tkeys:ctrl,s\n
explain:xxx\taction:MOVE_TO\tpoint:x,y\n
explain:xxx\taction:DRAG_TO\tpoint1:x1,y1\tpoint2:x2,y2\tbutton:left|right\n
explain:xxx\taction:DOUBLE_CLICK\tpoint:x,y\n
explain:xxx\taction:TRIPLE_CLICK\tpoint:x,y\n
explain:xxx\taction:MIDDLE_CLICK\tpoint:x,y\n
explain:xxx\taction:INFO\tvalue:xxxx\n
explain:xxx\taction:RIGHT_CLICK\tpoint:x,y\n
explain:xxx\taction:AWAKE\tvalue:xxxx\tpoint:x,y\n
"""

UITARS_USR_PROMPT_NOTHOUGHT = """You are a GUI agent. You are given a task and your action history, with screenshots. You need to perform the next action to complete the task. 
## Output Format
```
Action: ...
```
## Action Space
click(start_box='<|box_start|>(x1,y1)<|box_end|>')
left_double(start_box='<|box_start|>(x1,y1)<|box_end|>')
right_single(start_box='<|box_start|>(x1,y1)<|box_end|>')
drag(start_box='<|box_start|>(x1,y1)<|box_end|>', end_box='<|box_start|>(x3,y3)<|box_end|>')
hotkey(key='')
type(content='') #If you want to submit your input, use "\\n" at the end of `content`.
scroll(start_box='<|box_start|>(x1,y1)<|box_end|>', direction='down or up or right or left')
wait() #Sleep for 5s and take a screenshot to check for any changes.
finished()
call_user() # Submit the task and call the user when the task is unsolvable, or when you need the user's help.
## Coordinate System
All coordinates (x1, y1, x3, y3) should be normalized to the range [0, 1000], where:
- (0, 0) represents the top-left corner of the screen
- (1000, 1000) represents the bottom-right corner of the screen
- For example: click(start_box='<|box_start|>(500,500)<|box_end|>') clicks the center of the screen
## User Instruction
{instruction}
"""

UITARS_USR_PROMPT_THOUGHT = """你是一个【Computer屏幕操作专家】，用户将会给你看屏幕截图，你需要通过输出action 与用户的设备交互，从而完成用户的任务。

## 基本原则：
1. 你需要先思考后输出操作。思考过程以<THINK>开头，</THINK>结尾。内容格式：
    a. 你对屏幕的观察，屏幕里有什么？ 那些UI 激活了？ 排序情况如何？最高最低价等情况如何？当前界面是否在正确的道路上？... ?
    b. 用户的任务目前完成的如何了？还差什么？
    c. 你应该如何与界面中的UI 元素进行交互？
    d. 你最终决定怎么?

2. 坐标定义：屏幕坐标系以左上角为原点，x轴向右，y轴向下，范围为0到1000。
3. 每个action 中都会有explain 字段，用来简要解释你要做的事情。如 点击xxx； 打字输入内容等。
4. 每次可以输出多个action，用\n 分隔。

## 操作包含：
1. 点击（CLICK）屏幕上的坐标。需包含解释explain 解释你的点击意图；point表示点击的坐标。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:CLICK\tpoint:x,y\n
2. 输入（TYPE）一段文字。需包含，打字内容value。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:TYPE\tvalue:xxxx\tpoint:x,y\n
3. 完成（COMPLETE）任务。需包含return 表示向用户报告的内容。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:COMPLETE\tvalue:任务完成\n
4. 等待（WAIT）一段时间。需包含等待的时间value（秒）。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:WAIT\tvalue:n\n
5. 滑动（SCROLL）。滑动距离为value,坐标为point,需要在滑动点击目标区域。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:SCROLL\tvalue:n\tpoint:x,y\n
6. 滑动（HSCROLL）。滑动距离为value,坐标为point,需要在滑动点击目标区域。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:HSCROLL\tvalue:n\tpoint:x,y\n
7. 按键（HOTKEY）。按下或释放某个按键或组合键。需包含按键内容keys，keys中包含多个按键，用,分隔。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:HOTKEY\tkeys:ctrl,s\n
8. 鼠标移动（MOVE_TO）到某个坐标。需包含坐标point。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:MOVE_TO\tpoint:x,y\n
9. 拖拽（DRAG_TO）到某个坐标。需包含坐标point和鼠标button（left, right）。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:DRAG_TO\tpoint1:x1,y1\tpoint2:x2,y2\tbutton:left|right\n
10. 双击（DOUBLE_CLICK）。双击某个坐标。需包含坐标point。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:DOUBLE_CLICK\tpoint:x,y\n
11. 三击（TRIPLE_CLICK）。三击某个坐标。需包含坐标point。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:TRIPLE_CLICK\tpoint:x,y\n
12. 中间点击（MIDDLE_CLICK）。鼠标中键点击某个坐标。需包含坐标point。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:MIDDLE_CLICK\tpoint:x,y\n
13. 询问（INFO）。询问某个信息。需包含信息内容value。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:INFO\tvalue:xxxx\n
14. 右键点击（RIGHT_CLICK）。右键点击某个坐标。需包含坐标point。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:RIGHT_CLICK\tpoint:x,y\n
15. 唤醒（AWAKE）。唤醒某个app。需包含唤醒的app名称value。
例如：
<THINK> xxxxx </THINK>
explain:xxx\taction:AWAKE\tvalue:xxxx\tpoint:x,y\n


## 状态如下：

### 截图：
-----------用户的截图如下-----------
当前截图：
<image>
-----------用户的截图完成-----------

### 用户任务：
-----------用户任务如下-----------
{instruction}
-----------用户任务完成-----------

### 历史动作：
-----------历史动作如下-----------
{history_actions}
-----------历史动作完成-----------




## 指令：
请遵循指令，先思考，后输出对应的action，以<THINK> 开头【进行思考】，以</THINK>结束思考。
"""

FINISH_WORD = "finished"
WAIT_WORD = "wait"
ENV_FAIL_WORD = "error_env"
CALL_USER = "call_user"

IMAGE_FACTOR = 28
MIN_PIXELS = 100 * 28 * 28
MAX_PIXELS = 16384 * 28 * 28
MAX_RATIO = 200

def preprocess_box_coordinates(action_str):
    """
    预处理函数：将 <|box_start|>x y<|box_end|> 格式转换为 (x,y) 格式
    同时支持现有的 <|box_start|>(x,y)<|box_end|> 格式
    """
    import re
    
    # 模式1: <|box_start|>x y<|box_end|> (空格分隔的坐标)
    pattern1 = r'<\|box_start\|>(\d+)\s+(\d+)<\|box_end\|>'
    # 模式2: <|box_start|>(x,y)<|box_end|> (现有格式)
    pattern2 = r'<\|box_start\|>\((\d+),\s*(\d+)\)<\|box_end\|>'
    
    def replace_coords(match):
        x, y = match.groups()
        return f'({x},{y})'
    
    # 先处理空格分隔的格式
    action_str = re.sub(pattern1, replace_coords, action_str)
    # 再处理已有的括号格式（移除box标签）
    action_str = re.sub(pattern2, replace_coords, action_str)
    
    return action_str

# 定义一个函数来解析每个 action
def parse_action(action_str):
    """Parse an action string.

    First try key:value format (tab or space separated), then fallback to AST.
    """
    try:
        # New format: key:value pairs separated by tabs or spaces
        if '\t' in action_str or (' ' in action_str and ':' in action_str):
            delimiter = '\t' if '\t' in action_str else ' '

            if delimiter == ' ':
                import re as _re
                pattern = r'(\w+):([^\s]+(?:\s+[^\w:]*)*?)(?=\s+\w+:|$)'
                matches = _re.findall(pattern, action_str)
                parsed_action = {}
                action_type = None
                for key, value in matches:
                    value = value.strip()
                    if key == 'action':
                        action_type = value.strip().lower()
                    else:
                        parsed_action[key] = value
            else:
                parts = action_str.split(delimiter)
                parsed_action = {}
                action_type = None
                for part in parts:
                    if ':' in part:
                        key, value = part.split(':', 1)
                        if key == 'action':
                            action_type = value.strip().lower()
                        else:
                            parsed_action[key] = value

            if action_type:
                return {
                    'function': action_type,
                    'args': parsed_action
                }

        # Fallback to AST format
        action_str = preprocess_box_coordinates(action_str)
        node = ast.parse(action_str, mode='eval')
        if not isinstance(node, ast.Expression):
            raise ValueError("Not an expression")
        call = node.body
        if not isinstance(call, ast.Call):
            raise ValueError("Not a function call")
        if isinstance(call.func, ast.Name):
            func_name = call.func.id
        elif isinstance(call.func, ast.Attribute):
            func_name = call.func.attr
        else:
            func_name = None
        kwargs = {}
        for kw in call.keywords:
            key = kw.arg
            if isinstance(kw.value, ast.Constant):
                value = kw.value.value
            elif isinstance(kw.value, ast.Str):
                value = kw.value.s
            else:
                value = None
            kwargs[key] = value
        return {
            'function': func_name,
            'args': kwargs
        }
    except Exception as e:
        print(f"Failed to parse action '{action_str}': {e}")
        return None
    
def escape_single_quotes(text):
    # 匹配未转义的单引号（不匹配 \\'）
    pattern = r"(?<!\\)'"
    return re.sub(pattern, r"\\'", text)

def round_by_factor(number: int, factor: int) -> int:
    """Returns the closest integer to 'number' that is divisible by 'factor'."""
    return round(number / factor) * factor


def ceil_by_factor(number: int, factor: int) -> int:
    """Returns the smallest integer greater than or equal to 'number' that is divisible by 'factor'."""
    return math.ceil(number / factor) * factor


def floor_by_factor(number: int, factor: int) -> int:
    """Returns the largest integer less than or equal to 'number' that is divisible by 'factor'."""
    return math.floor(number / factor) * factor    


def linear_resize(
    height: int, width: int, factor: int = IMAGE_FACTOR, min_pixels: int = MIN_PIXELS, max_pixels: int = MAX_PIXELS
) -> tuple[int, int]:
    if width * height > max_pixels:
        """
        如果图片超过/低于像素限制，则计算一个缩放因子resize_factor，使图片的像素数缩小到等于或小于max_pixels。这个缩放因子是通过开平方根计算的，确保纵横比保持不变,这样原始的相对坐标可以不经转换直接复用
        """
        resize_factor = math.sqrt(max_pixels / (width * height))
        width, height = int(width * resize_factor), int(height * resize_factor)
    if width * height < min_pixels:
        resize_factor = math.sqrt(min_pixels / (width * height))
        width, height = math.ceil(width * resize_factor), math.ceil(height * resize_factor)

    return height, width 

def smart_resize(
    height: int, width: int, factor: int = IMAGE_FACTOR, min_pixels: int = MIN_PIXELS, max_pixels: int = MAX_PIXELS
) -> tuple[int, int]:
    """
    Rescales the image so that the following conditions are met:

    1. Both dimensions (height and width) are divisible by 'factor'.

    2. The total number of pixels is within the range ['min_pixels', 'max_pixels'].

    3. The aspect ratio of the image is maintained as closely as possible.
    """
    if max(height, width) / min(height, width) > MAX_RATIO:
        raise ValueError(
            f"absolute aspect ratio must be smaller than {MAX_RATIO}, got {max(height, width) / min(height, width)}"
        )
    h_bar = max(factor, round_by_factor(height, factor))
    w_bar = max(factor, round_by_factor(width, factor))
    if h_bar * w_bar > max_pixels:
        beta = math.sqrt((height * width) / max_pixels)
        h_bar = floor_by_factor(height / beta, factor)
        w_bar = floor_by_factor(width / beta, factor)
    elif h_bar * w_bar < min_pixels:
        beta = math.sqrt(min_pixels / (height * width))
        h_bar = ceil_by_factor(height * beta, factor)
        w_bar = ceil_by_factor(width * beta, factor)
    return h_bar, w_bar

def parse_action_to_structure_output(text, factor, origin_resized_height, origin_resized_width, model_type, max_pixels=16384*28*28, min_pixels=100*28*28):
    """Parse model output to structured action(s).

    This version aligns with stepcopilot_v1 to support the <THINK> ... </THINK>
    + multi-line key:value format, e.g. explain/action/point on separate lines.
    """
    text = text.strip()
    if model_type == "qwen25vl":
        smart_resize_height, smart_resize_width = smart_resize(
            origin_resized_height,
            origin_resized_width,
            factor=IMAGE_FACTOR,
            min_pixels=min_pixels,
            max_pixels=max_pixels,
        )

    # New format with <THINK> ... </THINK>
    thought = None
    if "<THINK>" in text and "</THINK>" in text:
        think_pattern = r"<THINK>\s*(.+?)\s*</THINK>"
        think_match = re.search(think_pattern, text, re.DOTALL)
        if think_match:
            thought = think_match.group(1).strip()
        # content after </THINK> is the action lines
        action_start = text.find("</THINK>")
        action_str = text[action_start + 8:].strip() if action_start != -1 else text
    else:
        # Backward compatibility: Thought/Action blocks
        if text.startswith("Thought:"):
            thought_pattern = r"Thought: (.+?)(?=\s*Action:|$)"
        elif text.startswith("Reflection:"):
            thought_pattern = r"Reflection: (.+?)Action_Summary: (.+?)(?=\s*Action:|$)"
        elif text.startswith("Action_Summary:"):
            thought_pattern = r"Action_Summary: (.+?)(?=\s*Action:|$)"
        else:
            thought_pattern = r"Thought: (.+?)(?=\s*Action:|$)"
        reflection = None
        thought_match = re.search(thought_pattern, text, re.DOTALL)
        if thought_match:
            if len(thought_match.groups()) == 1:
                thought = thought_match.group(1).strip()
            elif len(thought_match.groups()) == 2:
                thought = thought_match.group(2).strip()
                reflection = thought_match.group(1).strip()
        action_str = text.split("Action:")[-1] if "Action:" in text else text

    # Split into lines; allow a single action to be expressed over multiple key:value lines
    if "\n" in action_str:
        tmp_all_action = action_str.split("\n")
    else:
        tmp_all_action = [action_str]

    # Merge logically related lines into a single action line when they look like key:value pairs
    merged_actions = []
    current = []
    for line in tmp_all_action:
        s = line.strip()
        if not s:
            continue
        # If looks like a key:value line, keep accumulating
        if ":" in s and (s.startswith("explain:") or s.startswith("action:") or s.startswith("point:") or s.startswith("value:") or s.startswith("keys:") or s.startswith("direction:") or s.startswith("point1:") or s.startswith("point2:")):
            current.append(s)
        else:
            # not a key:value line; treat as a standalone
            if current:
                merged_actions.append("\t".join(current))
                current = []
            merged_actions.append(s)
    if current:
        merged_actions.append("\t".join(current))

    # Fallback to original action string if nothing merged
    if not merged_actions:
        merged_actions = [action_str]

    all_action = []
    for a in merged_actions:
        # If one string contains multiple actions, split into chunks per action
        if a.count("action:") > 1 or a.count("\taction:") > 1:
            # Tokenize key:value pairs by tab delimiter (preferred after merging)
            kv_pairs = [p for p in a.split("\t") if ":" in p]
            buckets = []
            cur = []
            saw_action = False
            for kv in kv_pairs:
                key = kv.split(":", 1)[0].strip().lower()
                if key == "action" and saw_action:
                    # starting a new action; flush previous
                    if cur:
                        buckets.append("\t".join(cur))
                    cur = [kv]
                else:
                    cur.append(kv)
                if key == "action":
                    saw_action = True
            if cur:
                buckets.append("\t".join(cur))
            # Drop any bucket that still lacks action (e.g., leading explain only)
            buckets = [b for b in buckets if "\taction:" in b or b.startswith("action:")]
            split_candidates = buckets if buckets else [a]
        else:
            split_candidates = [a]

        for cand in split_candidates:
            # Normalize TYPE old format if present
            if "type(content" in cand:
                def escape_quotes(match):
                    content = match.group(1)
                    return content
                pattern = r"type\(content='(.*?)'\)"
                content = re.sub(pattern, escape_quotes, cand)
                cand = escape_single_quotes(content)
                cand = "type(content='" + cand + "')"
            all_action.append(cand)

    parsed_actions = [parse_action(act.replace("\n", "\\n").lstrip()) for act in all_action]
    actions = []
    for action_instance, raw_str in zip(parsed_actions, all_action):
        if action_instance is None:
            # 忽略仅包含explain等信息但没有action指令的行，避免中断
            if raw_str.strip().startswith("explain:") and "action:" not in raw_str:
                continue
            print(f"Action can't parse: {raw_str}")
            raise ValueError(f"Action can't parse: {raw_str}")
        action_type = action_instance["function"]
        params = action_instance["args"]

        action_inputs = {}
        for param_name, param in params.items():
            if param == "":
                continue
            param = param.lstrip()

            # Support new format coordinates: point:x,y and drag point1/point2
            if param_name == "point":
                if "," in param:
                    # sanitize any stray chars like trailing 'e' from OCR/noise
                    import re as _re2
                    clean = _re2.sub(r"[^0-9,.-]", "", param)
                    x, y = clean.split(",")[:2]
                    float_x = float(x.strip()) / 1000.0
                    float_y = float(y.strip()) / 1000.0
                    action_inputs["start_box"] = str([float_x, float_y, float_x, float_y])
                else:
                    action_inputs[param_name.strip()] = param
            elif param_name in ["point1", "point2"]:
                if "," in param:
                    import re as _re3
                    clean = _re3.sub(r"[^0-9,.-]", "", param)
                    x, y = clean.split(",")[:2]
                    float_x = float(x.strip()) / 1000.0
                    float_y = float(y.strip()) / 1000.0
                    if param_name == "point1":
                        action_inputs["start_box"] = str([float_x, float_y, float_x, float_y])
                    else:
                        action_inputs["end_box"] = str([float_x, float_y, float_x, float_y])
                else:
                    action_inputs[param_name.strip()] = param
            elif param_name == "keys":
                action_inputs[param_name.strip()] = param
            elif "start_box" in param_name or "end_box" in param_name:
                # legacy box numbers
                ori_box = param
                numbers = ori_box.replace("(", "").replace(")", "").split(",")
                if model_type == "qwen25vl":
                    float_numbers = []
                    for num_idx, num in enumerate(numbers):
                        num = float(num)
                        if (num_idx + 1) % 2 == 0:
                            float_numbers.append(float(num / smart_resize_height))
                        else:
                            float_numbers.append(float(num / smart_resize_width))
                else:
                    float_numbers = [float(num) / factor for num in numbers]
                if len(float_numbers) == 2:
                    float_numbers = [float_numbers[0], float_numbers[1], float_numbers[0], float_numbers[1]]
                action_inputs[param_name.strip()] = str(float_numbers)
            else:
                action_inputs[param_name.strip()] = param

        actions.append({
            "reflection": None,
            "thought": thought,
            "action_type": action_type,
            "action_inputs": action_inputs,
            "text": text,
        })
    return actions

def parsing_response_to_pyautogui_code(responses, image_height: int, image_width:int, input_swap:bool=True) -> str:
    '''
    将M模型的输出解析为OSWorld中的action，生成pyautogui代码字符串
    参数:
        response: 包含模型输出的字典，结构类似于：
        {
            "action_type": "hotkey",
            "action_inputs": {
                "hotkey": "v ctrl",
                "start_box": None,
                "end_box": None
            }
        }
    返回:
        生成的pyautogui代码字符串
    '''

    pyautogui_code = f"import pyautogui\nimport time\n"
    if isinstance(responses, dict):
        responses = [responses]
    for response_id, response in enumerate(responses):
        if "observation" in response:
            observation = response["observation"]
        else:
            observation = ""

        if "thought" in response:
            thought = response["thought"]
        else:
            thought = ""
        
        if response_id == 0:
            pyautogui_code += f"'''\nObservation:\n{observation}\n\nThought:\n{thought}\n'''\n"
        else:
            pyautogui_code += f"\ntime.sleep(1)\n"

        action_dict = response
        action_type = action_dict.get("action_type")
        action_inputs = action_dict.get("action_inputs", {})
        
        if action_type == "hotkey":
            # Support both keys (comma-separated) and legacy key/hotkey
            if "keys" in action_inputs:
                hotkey = action_inputs.get("keys", "")
            elif "value" in action_inputs:
                hotkey = action_inputs.get("value", "")
            elif "key" in action_inputs:
                hotkey = action_inputs.get("key", "")
            else:
                hotkey = action_inputs.get("hotkey", "")

            if hotkey == "arrowleft":
                hotkey = "left"

            elif hotkey == "arrowright":
                hotkey = "right"
            
            elif hotkey == "arrowup":
                hotkey = "up"
            
            elif hotkey == "arrowdown":
                hotkey = "down"

            if hotkey:
                # Handle other hotkeys - support comma or space separated
                keys = hotkey.split(',') if ',' in hotkey else hotkey.split()
                convert_keys = []
                for key in keys:
                    key = key.strip()
                    if key == "space":
                        key = ' '
                    convert_keys.append(key)
                
                # 增强的热键执行 - 添加延迟和重试机制
                # pyautogui_code += f"\n# 执行热键: {hotkey}"
                # pyautogui_code += f"\n# 确保窗口获得焦点"
                # pyautogui_code += f"\npyautogui.click(960, 540)  # 点击屏幕中央确保焦点"
                # pyautogui_code += f"\ntime.sleep(0.5)  # 等待系统准备"
                # pyautogui_code += f"\ntry:"
                # pyautogui_code += f"\n    pyautogui.hotkey({', '.join([repr(k) for k in convert_keys])})"
                # pyautogui_code += f"\nexcept Exception as e:"
                # pyautogui_code += f"\n    print(f'热键失败，尝试备用方案: {{e}}')"
                # pyautogui_code += f"\n    # 备用方案：逐个按键"
                # for key in convert_keys:
                #     pyautogui_code += f"\n    pyautogui.keyDown({repr(key)})"
                # for key in reversed(convert_keys):
                #     pyautogui_code += f"\n    pyautogui.keyUp({repr(key)})"
                # pyautogui_code += f"\ntime.sleep(1.0)  # 等待热键响应"
        
        elif action_type == "press":
            # Parsing press action
            if "key" in action_inputs:
                key_to_press = action_inputs.get("key", "")
            else:
                key_to_press = action_inputs.get("press", "")

            if hotkey == "arrowleft":
                hotkey = "left"

            elif hotkey == "arrowright":
                hotkey = "right"
            
            elif hotkey == "arrowup":
                hotkey = "up"
            
            elif hotkey == "arrowdown":
                hotkey = "down"
            
            elif hotkey == "space":
                hotkey = " "
                
            if key_to_press:
                # Simulate pressing a single key
                pyautogui_code += f"\npyautogui.press({repr(key_to_press)})"
            
        elif action_type == "keyup":
            key_to_up = action_inputs.get("key", "")
            pyautogui_code += f"\npyautogui.keyUp({repr(key_to_up)})"
        
        elif action_type == "keydown":
            key_to_down = action_inputs.get("key", "")
            pyautogui_code += f"\npyautogui.keyDown({repr(key_to_down)})"

        elif action_type == "type":
            # Parsing typing action using clipboard
            content = action_inputs.get("content", "")
            content = escape_single_quotes(content)
            stripped_content = content
            if content.endswith("\n") or content.endswith("\\n"):
                stripped_content = stripped_content.rstrip("\\n").rstrip("\n")
            if content:
                if input_swap:
                    # 使用改进的输入方法替代原有的pyperclip方案
                    pyautogui_code += f"\n# 使用改进的输入方法"
                    pyautogui_code += f"\ntry:"
                    pyautogui_code += f"\n    from improved_input_methods import smart_type"
                    pyautogui_code += f"\n    success = smart_type('{stripped_content}')"
                    pyautogui_code += f"\n    if not success:"
                    pyautogui_code += f"\n        print('智能输入失败，尝试备用方案')"
                    pyautogui_code += f"\n        # 备用方案1: 尝试pyperclip"
                    pyautogui_code += f"\n        try:"
                    pyautogui_code += f"\n            import pyperclip"
                    pyautogui_code += f"\n            pyperclip.copy('{stripped_content}')"
                    pyautogui_code += f"\n            pyautogui.hotkey('ctrl', 'v')"
                    pyautogui_code += f"\n        except Exception as e:"
                    pyautogui_code += f"\n            print(f'pyperclip备用方案失败: {{e}}')"
                    pyautogui_code += f"\n            # 备用方案2: 直接输入"
                    pyautogui_code += f"\n            pyautogui.write('{stripped_content}', interval=0.05)"
                    pyautogui_code += f"\nexcept ImportError:"
                    pyautogui_code += f"\n    print('改进输入方法不可用，使用传统方案')"
                    pyautogui_code += f"\n    try:"
                    pyautogui_code += f"\n        import pyperclip"
                    pyautogui_code += f"\n        pyperclip.copy('{stripped_content}')"
                    pyautogui_code += f"\n        pyautogui.hotkey('ctrl', 'v')"
                    pyautogui_code += f"\n    except Exception as e:"
                    pyautogui_code += f"\n        print(f'传统pyperclip方案失败: {{e}}')"
                    pyautogui_code += f"\n        pyautogui.write('{stripped_content}', interval=0.05)"
                    pyautogui_code += f"\ntime.sleep(0.5)\n"
                    if content.endswith("\n") or content.endswith("\\n"):
                        pyautogui_code += f"\npyautogui.press('enter')"
                else:
                    pyautogui_code += f"\npyautogui.write('{stripped_content}', interval=0.05)"
                    pyautogui_code += f"\ntime.sleep(0.5)\n"
                    if content.endswith("\n") or content.endswith("\\n"):
                        pyautogui_code += f"\npyautogui.press('enter')"

        
        elif action_type in ["drag", "select"]:
            # Parsing drag or select action based on start and end_boxes
            start_box = action_inputs.get("start_box")
            end_box = action_inputs.get("end_box")
            if start_box and end_box:
                # Parse start box coordinates (relative 0-1)
                x1, y1, x2, y2 = eval(start_box)  # Assuming box is in [x1, y1, x2, y2]
                start_center_x = float((x1 + x2) / 2)
                start_center_y = float((y1 + y2) / 2)
                sx = round(start_center_x * image_width, 3)
                sy = round(start_center_y * image_height, 3)
                
                # Parse end box coordinates (relative 0-1)
                x1, y1, x2, y2 = eval(end_box)  # Assuming box is in [x1, y1, x2, y2]
                end_center_x = float((x1 + x2) / 2)
                end_center_y = float((y1 + y2) / 2)
                ex = round(end_center_x * image_width, 3)
                ey = round(end_center_y * image_height, 3)
                
                pyautogui_code += (
                    f"\npyautogui.moveTo({sx}, {sy})\n"
                    f"\npyautogui.dragTo({ex}, {ey}, duration=1.0)\n"
                )

        elif action_type == "scroll":
            # Parsing scroll action
            start_box = action_inputs.get("start_box")
            if start_box:
                # Parse scroll position coordinates (relative 0-1)
                x1, y1, x2, y2 = eval(start_box)  # Assuming box is in [x1, y1, x2, y2]
                scroll_center_x = float((x1 + x2) / 2)
                scroll_center_y = float((y1 + y2) / 2)
                x = round(scroll_center_x * image_width, 3)
                y = round(scroll_center_y * image_height, 3)
                
                # # 先点对应区域，再滚动
                # pyautogui_code += f"\npyautogui.click({x}, {y}, button='left')"
            else:
                x = None
                y = None
            direction = action_inputs.get("direction", "")
            value = action_inputs.get("value", None)
            
            if x == None:
                amount = 5
                try:
                    if value is not None:
                        amount = int(value)
                except:
                    amount = 5
                if "up" in direction.lower():
                    pyautogui_code += f"\npyautogui.scroll({amount})"
                elif "down" in direction.lower():
                    pyautogui_code += f"\npyautogui.scroll(-{amount})"
            else:
                amount = 5
                try:
                    if value is not None:
                        amount = int(value)
                except:
                    amount = 5
                if "up" in direction.lower():
                    pyautogui_code += f"\npyautogui.scroll({amount}, x={x}, y={y})"
                elif "down" in direction.lower():
                    pyautogui_code += f"\npyautogui.scroll(-{amount}, x={x}, y={y})"

        elif action_type == "hscroll":
            # Parsing horizontal scroll action
            start_box = action_inputs.get("start_box")
            if start_box:
                x1, y1, x2, y2 = eval(start_box)
                scroll_center_x = float((x1 + x2) / 2)
                scroll_center_y = float((y1 + y2) / 2)
                x = round(scroll_center_x * image_width, 3)
                y = round(scroll_center_y * image_height, 3)
            else:
                x = None
                y = None
            direction = action_inputs.get("direction", "")
            value = action_inputs.get("value", None)
            amount = 5
            try:
                if value is not None:
                    amount = int(value)
            except:
                amount = 5
            # pyautogui.hscroll: positive -> right, negative -> left
            if x is None:
                if "right" in direction.lower():
                    pyautogui_code += f"\npyautogui.hscroll({amount})"
                elif "left" in direction.lower():
                    pyautogui_code += f"\npyautogui.hscroll(-{amount})"
            else:
                if "right" in direction.lower():
                    pyautogui_code += f"\npyautogui.hscroll({amount}, x={x}, y={y})"
                elif "left" in direction.lower():
                    pyautogui_code += f"\npyautogui.hscroll(-{amount}, x={x}, y={y})"

        elif action_type in ["click", "left_single", "left_double", "right_single", "right_click", "hover", "double_click", "triple_click", "middle_click", "move_to", "drag_to"]:
            # Parsing mouse click actions
            start_box = action_inputs.get("start_box")
            # New format: if start_box not provided but point is, convert it
            if (not start_box or start_box == "None") and "point" in action_inputs:
                point = action_inputs.get("point", "")
                if "," in point:
                    x_str, y_str = point.split(",")
                    x = round(float(x_str.strip()) * image_width / 1000, 3)
                    y = round(float(y_str.strip()) * image_height / 1000, 3)
                else:
                    x, y = 0, 0
            else:
                start_box = str(start_box)
                if start_box:
                    start_box = eval(start_box)
                    if len(start_box) == 4:
                        x1, y1, x2, y2 = start_box
                    elif len(start_box) == 2:
                        x1, y1 = start_box
                        x2, y2 = x1, y1
                    center_x = float((x1 + x2) / 2)
                    center_y = float((y1 + y2) / 2)
                    x = round(center_x * image_width, 3)
                    y = round(center_y * image_height, 3)
                else:
                    x, y = 0, 0

            if action_type == "left_single" or action_type == "click":
                pyautogui_code += f"\npyautogui.click({x}, {y}, button='left')"
            elif action_type == "left_double" or action_type == "double_click":
                pyautogui_code += f"\npyautogui.doubleClick({x}, {y}, button='left')"
            elif action_type == "triple_click":
                pyautogui_code += f"\npyautogui.tripleClick({x}, {y})"
            elif action_type == "middle_click":
                pyautogui_code += f"\npyautogui.click({x}, {y}, button='middle')"
            elif action_type == "right_single":
                pyautogui_code += f"\npyautogui.click({x}, {y}, button='right')"
            elif action_type == "right_click":
                pyautogui_code += f"\npyautogui.rightClick({x}, {y})"
            elif action_type == "hover" or action_type == "move_to":
                pyautogui_code += f"\npyautogui.moveTo({x}, {y})"
            elif action_type == "drag_to":
                button = action_inputs.get("button", "left")
                pyautogui_code += f"\npyautogui.dragTo({x}, {y}, button='{button}', duration=1.0)"
        
        elif action_type in ["finished", "complete"]:
            pyautogui_code = f"DONE"
        
        else:
            pyautogui_code += f"\n# Unrecognized action type: {action_type}"

    return pyautogui_code

def add_box_token(input_string):
    # Step 1: Split the string into individual actions
    if "Action: " in input_string and "start_box=" in input_string:
        suffix = input_string.split("Action: ")[0] + "Action: "
        actions = input_string.split("Action: ")[1:]
        processed_actions = []
        for action in actions:
            action = action.strip()
            # Step 2: Extract coordinates (start_box or end_box) using regex
            coordinates = re.findall(r"(start_box|end_box)='\((\d+),\s*(\d+)\)'", action)
            
            updated_action = action  # Start with the original action
            for coord_type, x, y in coordinates:
                # Convert x and y to integers
                updated_action = updated_action.replace(f"{coord_type}='({x},{y})'", f"{coord_type}='<|box_start|>({x},{y})<|box_end|>'")
            processed_actions.append(updated_action)
        
        # Step 5: Reconstruct the final string
        final_string = suffix + "\n\n".join(processed_actions)
    else:
        final_string = input_string
    return final_string

def pil_to_base64(image):
    buffer = BytesIO()
    image.save(buffer, format="PNG")  # 你可以改成 "JPEG" 等格式
    return base64.b64encode(buffer.getvalue()).decode("utf-8")

def linearize_accessibility_tree(accessibility_tree, platform="ubuntu"):

    if platform == "ubuntu":
        _attributes_ns = attributes_ns_ubuntu
        _state_ns = state_ns_ubuntu
        _component_ns = component_ns_ubuntu
        _value_ns = value_ns_ubuntu
    elif platform == "windows":
        _attributes_ns = attributes_ns_windows
        _state_ns = state_ns_windows
        _component_ns = component_ns_windows
        _value_ns = value_ns_windows
    else:
        raise ValueError("Invalid platform, must be 'ubuntu' or 'windows'")

    filtered_nodes = filter_nodes(ET.fromstring(accessibility_tree), platform)
    linearized_accessibility_tree = [
        "tag\tname\ttext\tclass\tdescription\tposition (top-left x&y)\tsize (w&h)"
    ]

    # Linearize the accessibility tree nodes into a table format
    for node in filtered_nodes:
        if node.text:
            text = (
                node.text
                if '"' not in node.text
                else '"{:}"'.format(node.text.replace('"', '""'))
            )

        elif node.get("{{{:}}}class".format(class_ns_windows), "").endswith(
            "EditWrapper"
        ) and node.get("{{{:}}}value".format(_value_ns)):
            node_text = node.get("{{{:}}}value".format(_value_ns), "")
            text = (
                node_text
                if '"' not in node_text
                else '"{:}"'.format(node_text.replace('"', '""'))
            )
        else:
            text = '""'

        linearized_accessibility_tree.append(
            "{:}\t{:}\t{:}\t{:}\t{:}\t{:}\t{:}".format(
                node.tag,
                node.get("name", ""),
                text,
                (
                    node.get("{{{:}}}class".format(_attributes_ns), "")
                    if platform == "ubuntu"
                    else node.get("{{{:}}}class".format(class_ns_windows), "")
                ),
                node.get("{{{:}}}description".format(_attributes_ns), ""),
                node.get("{{{:}}}screencoord".format(_component_ns), ""),
                node.get("{{{:}}}size".format(_component_ns), ""),
            )
        )

    return "\n".join(linearized_accessibility_tree)

def trim_accessibility_tree(linearized_accessibility_tree, max_tokens):
    # enc = tiktoken.encoding_for_model("gpt-4")
    # tokens = enc.encode(linearized_accessibility_tree)
    # if len(tokens) > max_tokens:
    #     linearized_accessibility_tree = enc.decode(tokens[:max_tokens])
    #     linearized_accessibility_tree += "[...]\n"
    return linearized_accessibility_tree


class UITARSAgent:
    def __init__(
        self,
        model: str,
        runtime_conf: Dict,
        platform="ubuntu",
        action_space="pyautogui",
        observation_type="screenshot",
        # observation_type can be in ["screenshot", "a11y_tree", "screenshot_a11y_tree", "som"]
        max_trajectory_length=50,
        a11y_tree_max_tokens=10000,
        model_type="qwen25vl",
        **kwargs
    ):
        self.model = model
        self.platform = platform
        self.action_space = action_space
        self.observation_type = observation_type
        self.max_trajectory_length = max_trajectory_length
        self.a11y_tree_max_tokens = a11y_tree_max_tokens
        self.model_type = model_type
        self.runtime_conf = runtime_conf
        self.temperature = self.runtime_conf["temperature"]
        self.top_k = self.runtime_conf["top_k"]
        self.top_p = self.runtime_conf["top_p"]
        self.max_tokens = self.runtime_conf["max_tokens"]
        self.infer_mode = self.runtime_conf["infer_mode"]
        self.prompt_style = self.runtime_conf["prompt_style"]
        self.input_swap = self.runtime_conf["input_swap"]
        self.language = self.runtime_conf["language"]
        self.max_pixels = self.runtime_conf["max_pixels"]
        self.min_pixels = self.runtime_conf["min_pixels"]
        self.callusr_tolerance = self.runtime_conf["callusr_tolerance"]

        self.thoughts = []
        self.actions = []
        self.observations = []
        self.history_images = []
        self.history_responses = []
        
        self.prompt_action_space = STEPCOPILOT_ACTION_SPACE
        self.action_parse_res_factor = 1000
    
        self.prompt_template = UITARS_USR_PROMPT_THOUGHT
        
        if self.prompt_style == "qwen2vl_user" or self.prompt_style == "qwen25vl_normal":
            self.prompt_template = UITARS_USR_PROMPT_THOUGHT

        elif self.prompt_style == "qwen2vl_no_thought":
            self.prompt_template = UITARS_USR_PROMPT_NOTHOUGHT

        
        if "history_n" in self.runtime_conf:
            self.history_n = self.runtime_conf["history_n"]
        else:
            self.history_n = 5
        
        self.cur_callusr_count = 0

    def reset(self, runtime_logger=None):
        self.thoughts = []
        self.actions = []
        self.observations = []
        self.history_images = []
        self.history_responses = []
        

    def predict(
        self, instruction: str, obs: Dict, last_action_after_obs: Dict = None
    ) -> List:
        """
        Predict the next action(s) based on the current observation.
        """

        # Append trajectory
        # print(len(self.observations), len(self.actions), len(self.actions))
        assert len(self.observations) == len(self.actions) and len(self.actions) == len(
            self.thoughts
        ), "The number of observations and actions should be the same."

        if len(self.observations) > self.max_trajectory_length:
            if self.max_trajectory_length == 0:
                _observations = []
                _actions = []
                _thoughts = []
            else:
                _observations = self.observations[-self.max_trajectory_length :]
                _actions = self.actions[-self.max_trajectory_length :]
                _thoughts = self.thoughts[-self.max_trajectory_length :]
        else:
            _observations = self.observations
            _actions = self.actions
            _thoughts = self.thoughts


        self.history_images.append(obs["screenshot"])

        if self.observation_type in ["screenshot", "screenshot_a11y_tree"]:
            base64_image = obs["screenshot"]
            try:
                linearized_accessibility_tree = (
                    linearize_accessibility_tree(
                        accessibility_tree=obs["accessibility_tree"],
                        platform=self.platform,
                    )
                    if self.observation_type == "screenshot_a11y_tree"
                    else None
                )
            except:
                linearized_accessibility_tree = None
            # logger.debug("LINEAR AT: %s", linearized_accessibility_tree)

            if linearized_accessibility_tree:
                linearized_accessibility_tree = trim_accessibility_tree(
                    linearized_accessibility_tree, self.a11y_tree_max_tokens
                )

            if self.observation_type == "screenshot_a11y_tree":
                self.observations.append(
                    {
                        "screenshot": base64_image,
                        "accessibility_tree": linearized_accessibility_tree,
                    }
                )
            else:
                self.observations.append(
                    {"screenshot": base64_image, "accessibility_tree": None}
                )

        else:
            raise ValueError(
                "Invalid observation_type type: " + self.observation_type
            )  # 1}}}
        
        # 构建历史动作字符串
        history_actions_str = ""
        if len(self.history_responses) > 0:
            for history_idx, history_response in enumerate(self.history_responses):
                # 只获取最近的history_n个历史记录
                # if len(self.history_responses) - history_idx <= self.history_n:
                    # remove the content between <THINK> and </THINK>
                if "</THINK>" in history_response:
                    history_response = history_response.split("</THINK>")[1].strip()
                history_actions_str += f"idx={history_idx + 1}: {history_response}\n"
        
        # Use new Chinese prompt template
        user_prompt = self.prompt_template.format(
            instruction=instruction,
            history_actions=history_actions_str
        )
        # Fix last N images input bug
        # if len(self.history_images) > self.history_n:
        #     self.history_images = self.history_images[-self.history_n:]

        messages, images = [], []
        if isinstance(self.history_images, bytes):
            self.history_images = [self.history_images]
        elif isinstance(self.history_images, np.ndarray):
            self.history_images = list(self.history_images)
        elif isinstance(self.history_images, list):
            pass
        else:
            raise TypeError(f"Unidentified images type: {type(self.history_images)}")

        for turn, image in enumerate(self.history_images):
            # Fix last N images input bug
            # if len(images) >= self.history_n:
            #     break
            try:
                image = Image.open(BytesIO(image))
            except Exception as e:
                raise RuntimeError(f"Error opening image: {e}")

            # 固定图片分辨率为756x756
            #target_size = (756, 756)
            target_size = (1512, 1512)

            image = image.resize(target_size)

            if image.mode != "RGB":
                image = image.convert("RGB")

            images.append(image)
            # Debug: save the image locally
            # image.save(f"image_{turn}.png")

        messages = [
            {
                "role": "system",
                "content": "You are a helpful assistant."
            }
        ]
        
        # 构建符合step_modelproxy格式的消息
        content = []
        
        # 简化历史交互处理 - 历史动作已经在prompt中处理
        
        # 添加当前图片
        cur_image = images[-1]  # 使用最新的图片
        encoded_string = pil_to_base64(cur_image)
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{encoded_string}"}
        })
        
        # 添加用户提示
        content.append({
            "type": "text",
            "text": f"\n{user_prompt}"
        })
        
        # 添加用户消息
        messages.append({
            "role": "user",
            "content": content
        })

        try_times = 3
        origin_resized_height = images[-1].height
        origin_resized_width = images[-1].width
        temperature = self.temperature
        top_k = self.top_k
        while True:
            if try_times <= 0:
                print(f"Reach max retry times to fetch response from client, as error flag.")
                return "client error", ["DONE"]
            try:
                # 使用generate_response_v2函数替代OpenAI客户端调用
                # prediction = generate_response_v2(messages, self.model, temperature, self.max_tokens)
                prediction = ask_llm_anything("openai", self.model, messages, args={"max_tokens": self.max_tokens, "temperature": temperature})
                print("#" * 20)
                print("prediction:\n")
                print(prediction)
                print("\n#" * 20)
                
                if prediction:
                    print("*" * 20)
                    print("Response:")
                    print(prediction)
                    print("*" * 20)
                    prediction = prediction.strip()
                else:
                    prediction = None

            except Exception as e:
                logger.exception(f"Error when fetching response from client: {e}")
                prediction = None
                try_times -= 1
            
            try:
                parsed_responses = parse_action_to_structure_output(
                    prediction,
                    self.action_parse_res_factor,
                    origin_resized_height,
                    origin_resized_width,
                    self.model_type,
                    self.max_pixels,
                    self.min_pixels
                )
                break
            except Exception as e:
                print(f"Error when parsing response from client: {e}")
                # If fail to parse the model response, we use sampling parameters to avoid it
                prediction = None
                try_times -= 1
                temperature = 1
                top_k = -1
                
        if prediction is None:
            return "client error", ["DONE"]

        self.history_responses.append(prediction)
        self.thoughts.append(prediction)

        try:
            parsed_responses = parse_action_to_structure_output(
                prediction,
                self.action_parse_res_factor,
                origin_resized_height,
                origin_resized_width,
                self.model_type,
                self.max_pixels,
                self.min_pixels
            )
        except Exception as e:
            print(f"Parsing action error: {prediction}, with error:\n{e}")
            return f"Parsing action error: {prediction}, with error:\n{e}", ["DONE"]

        actions = []
        last_image = Image.open(BytesIO(self.history_images[-1]))
        obs_image_height = last_image.height
        obs_image_width = last_image.width
        for parsed_response in parsed_responses:
            if "action_type" in parsed_response:

                if parsed_response["action_type"] == FINISH_WORD:
                    self.actions.append(actions)

                    return prediction, ["DONE"]
                
                elif parsed_response["action_type"] == WAIT_WORD:
                    self.actions.append(actions)
                    return prediction, ["WAIT"]
                
                elif parsed_response["action_type"] == ENV_FAIL_WORD:
                    self.actions.append(actions)
                    return prediction, ["FAIL"]

                elif parsed_response["action_type"] == CALL_USER:
                    if self.callusr_tolerance > self.cur_callusr_count:
                        self.actions.append(actions)
                        self.cur_callusr_count += 1
                        return prediction, ["WAIT"]
                    else:
                        self.actions.append(actions)
                        return prediction, ["FAIL"]
            
            pyautogui_code = parsing_response_to_pyautogui_code(
                parsed_response,
                obs_image_height,
                obs_image_width,
                self.input_swap
            )
            actions.append(pyautogui_code)

        self.actions.append(actions)

        if len(self.history_responses) >= self.max_trajectory_length:
            # Default to FAIL if exceed max steps
            actions = ["FAIL"]

        return prediction, actions
