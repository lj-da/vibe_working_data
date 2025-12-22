"""
轨迹格式转换器：将 stepcopilot 格式转换为 cu_client 兼容格式
用于与 vis_traj.py 可视化工具集成
"""

import json
import os
import uuid
import datetime
import logging
from typing import Dict, List, Any
from megfile import smart_open, smart_makedirs
from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger(__name__)


class TrajectoryConverter:
    """轨迹格式转换器"""
    
    def __init__(self, s3_log_dir: str = "s3://tkj/os-copilot-local-eval-logs/traces", 
                 s3_image_dir: str = "s3://tkj/os-copilot-local-eval-logs/images"):
        self.s3_log_dir = s3_log_dir
        self.s3_image_dir = s3_image_dir
        
    def generate_session_id(self) -> str:
        """生成与 cu_client 兼容的 session_id"""
        return str(uuid.uuid4())
    
    def convert_stepcopilot_to_cu_client_format(self, 
                                              stepcopilot_logs: List[Dict], 
                                              session_id: str,
                                              task: str,
                                              model_config: Dict,
                                              domain: str = "unknown",
                                              example_id: str = "unknown") -> List[Dict]:
        """
        将 stepcopilot 格式的轨迹转换为 cu_client 格式
        
        Args:
            stepcopilot_logs: stepcopilot 格式的轨迹日志列表
            session_id: 会话ID
            task: 任务描述
            model_config: 模型配置
            domain: 任务域
            example_id: 示例ID
            
        Returns:
            cu_client 格式的轨迹日志列表
        """
        cu_client_logs = []
        
        # 1. 添加 session_start 配置日志
        config_log = {
            "session_id": session_id,
            "timestamp": datetime.datetime.now().isoformat(),
            "message": {
                "log_type": "session_start",
                "task": task,
                "model_config": model_config,
                "domain": domain,
                "example_id": example_id,
                "trajectory_format": "stepcopilot_converted"
            }
        }
        cu_client_logs.append(config_log)
        
        # 2. 转换每个步骤的日志
        for step_log in stepcopilot_logs:
            if "Error" in step_log:
                # 错误日志
                error_log = {
                    "session_id": session_id,
                    "timestamp": step_log.get("action_timestamp", datetime.datetime.now().isoformat()),
                    "message": {
                        "log_type": "error",
                        "error": step_log["Error"],
                        "step_num": step_log.get("step_num", 0)
                    }
                }
                cu_client_logs.append(error_log)
                continue
            
            # 正常步骤日志
            step_num = step_log.get("step_num", 0)
            action_timestamp = step_log.get("action_timestamp", "")
            
            # 构建 environment 信息
            environment = {
                "image": f"{self.s3_image_dir}/{session_id}_step_{step_num}_{action_timestamp}.jpeg",
                "step_num": step_num,
                "timestamp": action_timestamp,
                "reward": step_log.get("reward", 0.0),
                "done": step_log.get("done", False),
                "info": step_log.get("info", {})
            }
            
            # 构建 action 信息
            action = step_log.get("action", {})
            response = step_log.get("response", "")
            
            # 解析 response 中的动作信息
            parsed_action = self._parse_response_action(response)
            
            # 如果 action 是字符串（Python代码），保存为 raw_action
            if isinstance(action, str):
                action_dict = {
                    "raw_action": action,
                    "explain": parsed_action.get("explain", f"Step {step_num} action"),
                    "action": parsed_action.get("action", "UNKNOWN")
                }
                
                # 根据动作类型只添加相关参数
                action_type = parsed_action.get("action", "UNKNOWN")
                self._add_relevant_params(action_dict, parsed_action, action_type)
            else:
                # 如果 action 已经是字典，合并解析的信息
                action_dict = action.copy()
                if "explain" not in action_dict:
                    action_dict["explain"] = parsed_action.get("explain", f"Step {step_num} action")
                if "action" not in action_dict and "action_type" not in action_dict:
                    action_dict["action"] = parsed_action.get("action", "UNKNOWN")
                
                # 根据动作类型只添加相关参数
                action_type = parsed_action.get("action", "UNKNOWN")
                self._add_relevant_params(action_dict, parsed_action, action_type)
            
            # 添加思考过程（如果有 response）
            if response:
                action_dict["cot"] = response
            
            action = action_dict
            
            # 构建完整的日志条目
            step_log_entry = {
                "session_id": session_id,
                "timestamp": action_timestamp,
                "message": {
                    "environment": environment,
                    "action": action,
                    "step_num": step_num,
                    "llm_cost": {
                        "llm_time": 0.0,  # 默认值，实际应该从 agent 获取
                        "llm_start_time": 0,
                        "llm_end_time": 0
                    }
                }
            }
            
            cu_client_logs.append(step_log_entry)
        
        return cu_client_logs
    
    def _parse_response_action(self, response: str) -> Dict:
        """
        解析 response 中的动作信息
        
        Args:
            response: StepCopilot 的 response 字符串
            
        Returns:
            解析后的动作字典
        """
        parsed = {
            "action": "UNKNOWN",
            "point": None,
            "value": None,
            "keys": None,
            "point1": None,
            "point2": None,
            "button": None,
            "explain": ""
        }
        
        if not response:
            return parsed
        
        try:
            import re
            
            # 提取 action 类型
            action_match = re.search(r'action:(\w+)', response)
            if action_match:
                parsed["action"] = action_match.group(1)
            
            # 提取 explain 信息（在 action: 之前的部分）
            explain_match = re.search(r'explain:(.*?)\taction:', response)
            if explain_match:
                parsed["explain"] = explain_match.group(1).strip()
            else:
                # 如果没有找到 explain，尝试提取整个 response 作为解释
                parsed["explain"] = response.strip()
            
            # 根据动作类型提取相应的参数
            action_type = parsed["action"]
            
            if action_type in ["CLICK", "RIGHT_CLICK", "DOUBLE_CLICK", "TRIPLE_CLICK", "MIDDLE_CLICK", "MOVE_TO", "TYPE", "SCROLL", "HSCROLL", "AWAKE"]:
                # 这些动作需要 point 参数
                point_match = re.search(r'point:(\d+),(\d+)', response)
                if point_match:
                    parsed["point"] = [int(point_match.group(1)), int(point_match.group(2))]
            
            if action_type in ["TYPE", "COMPLETE", "WAIT", "SCROLL", "HSCROLL", "INFO", "AWAKE"]:
                # 这些动作需要 value 参数
                value_match = re.search(r'value:([^\t\n]+)', response)
                if value_match:
                    parsed["value"] = value_match.group(1).strip()
            
            if action_type == "HOTKEY":
                # HOTKEY 动作需要 keys 参数
                keys_match = re.search(r'keys:([^\t\n]+)', response)
                if keys_match:
                    parsed["keys"] = keys_match.group(1).strip()
            
            if action_type == "DRAG_TO":
                # DRAG_TO 动作需要 point1, point2 和 button 参数
                point1_match = re.search(r'point1:(\d+),(\d+)', response)
                if point1_match:
                    parsed["point1"] = [int(point1_match.group(1)), int(point1_match.group(2))]
                
                point2_match = re.search(r'point2:(\d+),(\d+)', response)
                if point2_match:
                    parsed["point2"] = [int(point2_match.group(1)), int(point2_match.group(2))]
                
                button_match = re.search(r'button:([^\t\n]+)', response)
                if button_match:
                    parsed["button"] = button_match.group(1).strip()
                
        except Exception as e:
            logger.warning(f"解析 response 动作信息失败: {e}")
            parsed["explain"] = response.strip()
        
        return parsed
    
    def _maybe_draw_action_overlay(self, image_path: str, action: Dict, step_num: int) -> str:
        """当动作包含坐标时，在截图上绘制标记并返回新文件路径；否则返回空字符串。
        支持 CLICK/DOUBLE_CLICK/TRIPLE_CLICK/RIGHT_CLICK/MIDDLE_CLICK/MOVE_TO/DRAG_TO/SCROLL/HSCROLL/TYPE（当带 point 时）。
        """
        try:
            if not isinstance(action, dict):
                return ""
            action_type = (action.get("action") or action.get("action_type") or "").upper()
            if action_type not in [
                "CLICK", "DOUBLE_CLICK", "TRIPLE_CLICK", "RIGHT_CLICK", "MIDDLE_CLICK", "MOVE_TO", "DRAG_TO", "SCROLL", "HSCROLL", "TYPE"
            ]:
                return ""

            # 选择坐标：DRAG_TO 使用 point2；否则使用 point
            point = None
            if action_type == "DRAG_TO" and action.get("point2"):
                point = action.get("point2")
            else:
                point = action.get("point")
            if not point or not isinstance(point, (list, tuple)) or len(point) < 2:
                return ""

            x, y = int(point[0]), int(point[1])

            # 打开图片并绘制
            image = Image.open(image_path).convert('RGBA')
            overlay = Image.new('RGBA', image.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)

            color_map = {
                'CLICK': (255, 0, 0, 140),
                'DOUBLE_CLICK': (255, 128, 0, 140),
                'TRIPLE_CLICK': (255, 200, 0, 140),
                'RIGHT_CLICK': (0, 200, 0, 140),
                'MIDDLE_CLICK': (0, 0, 255, 140),
                'MOVE_TO': (255, 0, 255, 120),
                'DRAG_TO': (0, 255, 255, 140),
                'SCROLL': (180, 180, 180, 140),
                'HSCROLL': (150, 150, 255, 140),
                'TYPE': (200, 0, 200, 140),
            }
            color = color_map.get(action_type, (255, 0, 0, 140))
            radius = 16
            draw.ellipse([x - radius, y - radius, x + radius, y + radius], fill=color)

            label = f"{step_num}.{action_type}"
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 14)
            except Exception:
                font = ImageFont.load_default()
            bbox = draw.textbbox((0, 0), label, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            tx, ty = x + radius + 6, y - radius - th - 6
            if tx + tw > image.size[0]:
                tx = x - radius - tw - 6
            if ty < 0:
                ty = y + radius + 6
            draw.rectangle([tx - 2, ty - 2, tx + tw + 2, ty + th + 2], fill=(0, 0, 0, 180))
            draw.text((tx, ty), label, fill=(255, 255, 255, 255), font=font)

            composed = Image.alpha_composite(image, overlay).convert('RGB')
            out_path = image_path.replace('.png', f'_overlay_step{step_num}.png')
            composed.save(out_path)
            return out_path
        except Exception as e:
            logger.warning(f"生成点击可视化失败（使用原图上传）: {e}")
            return ""
    
    def _add_relevant_params(self, action_dict: Dict, parsed_action: Dict, action_type: str):
        """
        根据动作类型只添加相关参数
        
        Args:
            action_dict: 要更新的动作字典
            parsed_action: 解析出的动作参数
            action_type: 动作类型
        """
        # 根据动作类型添加相关参数
        if action_type in ["CLICK", "RIGHT_CLICK", "DOUBLE_CLICK", "TRIPLE_CLICK", "MIDDLE_CLICK", "MOVE_TO"]:
            # 只需要 point 参数
            if parsed_action.get("point"):
                action_dict["point"] = parsed_action["point"]
        
        elif action_type in ["TYPE"]:
            # 需要 point 和 value 参数
            if parsed_action.get("point"):
                action_dict["point"] = parsed_action["point"]
            if parsed_action.get("value"):
                action_dict["value"] = parsed_action["value"]
        
        elif action_type in ["SCROLL", "HSCROLL"]:
            # 需要 point 和 value 参数
            if parsed_action.get("point"):
                action_dict["point"] = parsed_action["point"]
            if parsed_action.get("value"):
                action_dict["value"] = parsed_action["value"]
        
        elif action_type == "HOTKEY":
            # 只需要 keys 参数
            if parsed_action.get("keys"):
                action_dict["keys"] = parsed_action["keys"]
        
        elif action_type == "DRAG_TO":
            # 需要 point1, point2 和 button 参数
            if parsed_action.get("point1"):
                action_dict["point1"] = parsed_action["point1"]
            if parsed_action.get("point2"):
                action_dict["point2"] = parsed_action["point2"]
            if parsed_action.get("button"):
                action_dict["button"] = parsed_action["button"]
        
        elif action_type in ["COMPLETE", "WAIT", "INFO"]:
            # 只需要 value 参数
            if parsed_action.get("value"):
                action_dict["value"] = parsed_action["value"]
        
        elif action_type == "AWAKE":
            # 需要 point 和 value 参数
            if parsed_action.get("point"):
                action_dict["point"] = parsed_action["point"]
            if parsed_action.get("value"):
                action_dict["value"] = parsed_action["value"]
    
    def upload_trajectory_to_s3(self, cu_client_logs: List[Dict], session_id: str) -> str:
        """
        将转换后的轨迹上传到 S3
        
        Args:
            cu_client_logs: cu_client 格式的轨迹日志
            session_id: 会话ID
            
        Returns:
            S3 轨迹文件路径
        """
        s3_trajectory_path = f"{self.s3_log_dir}/{session_id}.jsonl"
        
        try:
            # 对于 S3 路径，不需要创建目录，直接写入文件
            # smart_makedirs 对 S3 路径处理有问题
            
            # 写入轨迹文件
            with smart_open(s3_trajectory_path, "w") as f:
                for log_entry in cu_client_logs:
                    f.write(json.dumps(log_entry, ensure_ascii=False) + "\n")
            
            logger.info(f"✅ 轨迹已上传到 S3: {s3_trajectory_path}")
            return s3_trajectory_path
            
        except Exception as e:
            logger.error(f"❌ 上传轨迹到 S3 失败: {e}")
            raise
    
    def upload_screenshot_to_s3(self, local_image_path: str, session_id: str, step_num: int, timestamp: str) -> str:
        """
        将截图上传到 S3
        
        Args:
            local_image_path: 本地截图路径
            session_id: 会话ID
            step_num: 步骤编号
            timestamp: 时间戳
            
        Returns:
            S3 截图文件路径
        """
        s3_image_path = f"{self.s3_image_dir}/{session_id}_step_{step_num}_{timestamp}.jpeg"
        
        try:
            # 对于 S3 路径，不需要创建目录，直接上传文件
            
            # 读取本地图片并上传
            with open(local_image_path, "rb") as local_f:
                image_data = local_f.read()
            
            with smart_open(s3_image_path, "wb") as s3_f:
                s3_f.write(image_data)
            
            logger.info(f"✅ 截图已上传到 S3: {s3_image_path}")
            return s3_image_path
            
        except Exception as e:
            logger.error(f"❌ 上传截图到 S3 失败: {e}")
            raise
    
    def convert_and_upload_trajectory(self, 
                                    stepcopilot_logs: List[Dict],
                                    session_id: str,
                                    task: str,
                                    model_config: Dict,
                                    domain: str = "unknown",
                                    example_id: str = "unknown",
                                    local_screenshots_dir: str = None) -> str:
        """
        完整的转换和上传流程
        
        Args:
            stepcopilot_logs: stepcopilot 格式的轨迹日志
            session_id: 会话ID
            task: 任务描述
            model_config: 模型配置
            domain: 任务域
            example_id: 示例ID
            local_screenshots_dir: 本地截图目录
            
        Returns:
            S3 轨迹文件路径
        """
        # 1. 转换格式
        cu_client_logs = self.convert_stepcopilot_to_cu_client_format(
            stepcopilot_logs, session_id, task, model_config, domain, example_id
        )
        
        # 2. 上传截图（如果提供了本地截图目录）
        if local_screenshots_dir and os.path.exists(local_screenshots_dir):
            for log_entry in cu_client_logs:
                if "environment" in log_entry["message"]:
                    env = log_entry["message"]["environment"]
                    action_for_step = log_entry["message"].get("action", {})
                    if "step_num" in env:
                        step_num = env["step_num"]
                        timestamp = env.get("timestamp", "")
                        
                        # 查找对应的本地截图文件
                        local_image_pattern = f"step_{step_num}_{timestamp}.png"
                        local_image_path = os.path.join(local_screenshots_dir, local_image_pattern)
                        
                        if os.path.exists(local_image_path) and os.path.getsize(local_image_path) > 0:
                            # 若为点击/移动/拖拽等包含坐标的动作，先在图上绘制可视化点
                            overlay_image_path = self._maybe_draw_action_overlay(local_image_path, action_for_step, step_num)
                            try:
                                s3_image_path = self.upload_screenshot_to_s3(
                                    overlay_image_path or local_image_path, session_id, step_num, timestamp
                                )
                                # 更新环境中的图片路径
                                env["image"] = s3_image_path
                            except Exception as e:
                                logger.warning(f"⚠️ 上传截图失败: {e}")
                                # 如果上传失败，使用默认的 S3 路径
                                env["image"] = f"{self.s3_image_dir}/{session_id}_step_{step_num}_{timestamp}.jpeg"
                        else:
                            # 如果本地图片不存在或为空，使用默认的 S3 路径
                            env["image"] = f"{self.s3_image_dir}/{session_id}_step_{step_num}_{timestamp}.jpeg"
        
        # 3. 上传轨迹
        s3_trajectory_path = self.upload_trajectory_to_s3(cu_client_logs, session_id)
        
        return s3_trajectory_path


def create_model_config_from_args(args) -> Dict:
    """从命令行参数创建模型配置"""
    return {
        "model_name": args.model,
        "model_type": args.model_type,
        "infer_mode": args.infer_mode,
        "prompt_style": args.prompt_style,
        "language": args.language,
        "temperature": args.temperature,
        "top_p": args.top_p,
        "top_k": args.top_k,
        "max_tokens": args.max_tokens,
        "max_pixels": args.max_pixels,
        "min_pixels": args.min_pixels,
        "history_n": args.history_n,
        "callusr_tolerance": args.callusr_tolerance
    }
