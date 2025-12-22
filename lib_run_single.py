import datetime
import json
import logging
import os
import time
from wrapt_timeout_decorator import *
from trajectory_converter import TrajectoryConverter, create_model_config_from_args
from session_id_manager import SessionIDManager

logger = logging.getLogger("desktopenv.experiment")


def run_single_example(agent, env, example, max_steps, instruction, args, example_result_dir, scores, session_id_manager=None):
    runtime_logger = setup_logger(example, example_result_dir)
    try:
        agent.reset(runtime_logger)
    except Exception as e:
        agent.reset()

    env.reset(task_config=example)
    
    time.sleep(60) # Wait for the environment to be ready
    obs = env._get_obs() # Get the initial observation
    done = False
    step_idx = 0
    
    # 初始化轨迹转换器
    converter = TrajectoryConverter()
    session_id = converter.generate_session_id()
    logger.info(f"🆔 生成会话ID: {session_id}")
    
    # 创建模型配置
    model_config = create_model_config_from_args(args)
    
    # 存储轨迹日志用于后续转换
    stepcopilot_logs = []
    
    # 记录任务开始信息
    example_id = example.get('id', 'unknown')
    domain = getattr(args, 'domain', 'unknown')
    
    env.controller.start_recording()
    while not done and step_idx < max_steps:
        response, actions = agent.predict(
            instruction,
            obs
        )
        for action in actions:
            # Capture the timestamp before executing the action
            action_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
            logger.info("Step %d: %s", step_idx + 1, action)
            obs, reward, done, info = env.step(action, args.sleep_after_execution)

            logger.info("Reward: %.2f", reward)
            logger.info("Done: %s", done)
            
            # Save screenshot locally
            screenshot_filename = f"step_{step_idx + 1}_{action_timestamp}.png"
            screenshot_path = os.path.join(example_result_dir, screenshot_filename)
            with open(screenshot_path, "wb") as _f:
                _f.write(obs['screenshot'])
            
            # 构建步骤日志
            step_log = {
                "step_num": step_idx + 1,
                "action_timestamp": action_timestamp,
                "action": action,
                "response": response,
                "reward": reward,
                "done": done,
                "info": info,
                "screenshot_file": screenshot_filename
            }
            
            # 保存到本地轨迹文件
            with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                f.write(json.dumps(step_log))
                f.write("\n")
            
            # 添加到内存中的日志列表
            stepcopilot_logs.append(step_log)
            
            if done:
                logger.info("The episode is done.")
                break
        step_idx += 1
    
    result = env.evaluate()
    logger.info("Result: %.2f", result)
    scores.append(result)
    
    # 确定停止原因
    stop_reason = "completed" if done else "max_steps_reached"
    
    # 保存结果文件
    with open(os.path.join(example_result_dir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{result}\n")
    
    # 停止录制
    env.controller.end_recording(os.path.join(example_result_dir, "recording.mp4"))
    
    # 转换并上传轨迹到 S3（如果启用了 S3 上传）
    s3_upload_success = False
    if getattr(args, 'upload_to_s3', False):
        try:
            logger.info("🔄 开始转换并上传轨迹到 S3...")
            s3_trajectory_path = converter.convert_and_upload_trajectory(
                stepcopilot_logs=stepcopilot_logs,
                session_id=session_id,
                task=instruction,
                model_config=model_config,
                domain=getattr(args, 'domain', 'unknown'),
                example_id=example.get('id', 'unknown'),
                local_screenshots_dir=example_result_dir
            )
            logger.info(f"✅ 轨迹已成功上传到 S3: {s3_trajectory_path}")
            logger.info(f"🔍 可使用 vis_traj.py 查看轨迹，Session ID: {session_id}")
            s3_upload_success = True
            
            # 保存 session_id 到本地文件
            with open(os.path.join(example_result_dir, "session_id.txt"), "w") as f:
                f.write(session_id)
                
        except Exception as e:
            logger.error(f"❌ 上传轨迹到 S3 失败: {e}")
            logger.warning("⚠️ 轨迹仅保存在本地，无法使用 vis_traj.py 查看")
    else:
        logger.info(f"📝 轨迹已保存到本地: {example_result_dir}/traj.jsonl")
        logger.info(f"🆔 Session ID: {session_id} (未上传到 S3)")
    
    # 记录 Session ID 到汇总文件
    if session_id_manager:
        try:
            additional_info = {
                "instruction": instruction,
                "s3_upload_success": s3_upload_success,
                "upload_to_s3_enabled": getattr(args, 'upload_to_s3', False),
                "result_dir": example_result_dir
            }
            
            session_id_manager.add_session_id(
                session_id=session_id,
                example_id=example_id,
                domain=domain,
                result=result,
                stop_reason=stop_reason,
                steps=step_idx + 1,
                additional_info=additional_info
            )
            
            logger.info(f"📝 Session ID 已记录到汇总文件: {session_id}")
            
        except Exception as e:
            logger.error(f"❌ 记录 Session ID 失败: {e}")
    
    return session_id


def setup_logger(example, example_result_dir):
    runtime_logger = logging.getLogger(f"desktopenv.example.{example['id']}")
    runtime_logger.setLevel(logging.DEBUG)
    runtime_logger.addHandler(logging.FileHandler(os.path.join(example_result_dir, "runtime.log")))
    return runtime_logger

def run_single_example_human(env, example, max_steps, instruction, args, example_result_dir, scores):
    """人工操作模式：显示任务并等待用户完成后手动验证"""
    runtime_logger = setup_logger(example, example_result_dir)
    
    print("\n" + "="*80)
    print("🎯 新任务开始 / New Task Started")
    print("="*80)
    print(f"📝 任务描述 / Task Instruction: {instruction}")
    print(f"📂 示例ID / Example ID: {example.get('id', 'Unknown')}")
    print(f"🏷️  应用类型 / Application: {example.get('app', 'Unknown')}")
    print("="*80)
    
    env.reset(task_config=example)
    env.controller.start_recording()
    
    print("⏳ 等待环境准备就绪... / Waiting for environment to be ready...")
    time.sleep(5)  # 减少等待时间
    
    obs = env._get_obs() # Get the initial observation
    
    # Save initial screenshot
    action_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
    with open(os.path.join(example_result_dir, f"initial_state_{action_timestamp}.png"), "wb") as _f:
        _f.write(obs['screenshot'])
    
    # Save trajectory information
    with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
        f.write(json.dumps({
            "instruction": instruction,
            "initial_state": f"initial_state_{action_timestamp}.png",
            "start_time": action_timestamp,
            "mode": "manual_operation"
        }))
        f.write("\n")
    
    print("\n🖥️  环境已准备就绪！/ Environment is ready!")
    print("📋 请根据上述任务描述在虚拟机中进行操作")
    print("📋 Please perform the task according to the instruction above")
    print("\n💡 操作提示 / Operation Tips:")
    print("   - 请在虚拟机窗口中完成所需的操作")
    print("   - 完成后，请返回此命令行窗口")
    print("   - Please complete the required operations in the VM window")
    print("   - After completion, return to this command line window")
    
    print("\n" + "-"*60)
    print("⌨️  完成任务后，请按回车键继续... / Press Enter after completing the task...")
    print("-"*60)
    
    # 等待用户按回车键
    input()
    
    print("\n📊 正在评估任务完成情况... / Evaluating task completion...")
    
    # 获取最终状态
    final_obs = env._get_obs()
    final_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
    
    # 保存最终截图
    with open(os.path.join(example_result_dir, f"final_state_{final_timestamp}.png"), "wb") as _f:
        _f.write(final_obs['screenshot'])
    
    # 评估结果
    result = env.evaluate()
    
    print(f"\n📈 评估结果 / Evaluation Result: {result:.2f}")
    
    if result >= 1.0:
        print("✅ 任务成功完成！/ Task completed successfully!")
    elif result >= 0.5:
        print("⚠️ 任务部分完成 / Task partially completed")
    else:
        print("❌ 任务未完成 / Task not completed")
    
    logger.info("Human operation result: %.2f", result)
    scores.append(result)
    
    # 保存最终轨迹信息
    with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
        f.write(json.dumps({
            "final_state": f"final_state_{final_timestamp}.png",
            "end_time": final_timestamp,
            "result": result,
            "mode": "manual_operation",
            "evaluation": "human_completed"
        }))
        f.write("\n")
    
    # 保存结果文件
    with open(os.path.join(example_result_dir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{result}\n")
    
    # 停止录制
    env.controller.end_recording(os.path.join(example_result_dir, "recording.mp4"))
    
    print("🎬 操作录制已保存 / Operation recording saved")
    print("="*80)



def run_single_example_openaicua(agent, env, example, max_steps, instruction, args, example_result_dir, scores):
    runtime_logger = setup_logger(example, example_result_dir)
    agent.reset(runtime_logger)
    env.reset(task_config=example)
    time.sleep(60) # Wait for the environment to be ready
    obs = env._get_obs() # Get the initial observation
    done = False
    step_idx = 0
    env.controller.start_recording()
    while not done and step_idx < max_steps:
        response, actions = agent.predict(
            instruction,
            obs
        )

        done = not response.get('state_correct', False)

        for action in actions:
            # Capture the timestamp before executing the action
            action_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
            logger.info("Step %d: %s", step_idx + 1, action)
            obs, reward, done, info, step_info = agent.step(action)

            if not done:
                if not response.get('state_correct', False):
                    done = True

            logger.info("Reward: %.2f", reward)
            logger.info("Done: %s", done)
            # Save screenshot and trajectory information
            with open(os.path.join(example_result_dir, f"step_{step_idx + 1}_{action_timestamp}.png"),
                      "wb") as _f:
                _f.write(obs['screenshot'])

            # Remove pending checks if they exist which will cause issues with json serialization
            if action.get('pending_checks', None):
                del action['pending_checks']

            with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                f.write(json.dumps({
                    "step_num": step_idx + 1,
                    "action_timestamp": action_timestamp,
                    "action": action,
                    "reward": reward,
                    "done": done,
                    "info": info,
                    "screenshot_file": f"step_{step_idx + 1}_{action_timestamp}.png"
                }))
                f.write("\n")
            if done:
                logger.info("The episode is done.")
                break
        step_idx += 1
    result = env.evaluate()
    logger.info("Result: %.2f", result)
    scores.append(result)
    with open(os.path.join(example_result_dir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{result}\n")
    env.controller.end_recording(os.path.join(example_result_dir, "recording.mp4"))

def run_single_example_opencua(agent, env, example, max_steps, instruction, args, example_result_dir, scores):
    runtime_logger = setup_logger(example, example_result_dir)
    agent.reset(runtime_logger)
    env.reset(task_config=example)
    time.sleep(60) # Wait for the environment to be ready
    obs = env._get_obs() # Get the initial observation
    done = False
    step_idx = 0
    env.controller.start_recording()
    while not done and step_idx < max_steps:
        response, actions, info_dict = agent.predict(instruction, obs)

        logger.info(f"Got Action: {actions}")
        # Breack if no actions
        if not actions or len(actions)==0 or actions[0]=="" or actions[0].lower().startswith("error"): 
            break

        for action in actions:
            # Capture the timestamp before executing the action
            action_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
            logger.info("Step %d: %s", step_idx + 1, action)
            
            obs, reward, done, info = env.step(action, args.sleep_after_execution)

            logger.info(f"Action {action} executed, reward: {reward}, done: {done}")
            # Save screenshot and trajectory information
            with open(os.path.join(example_result_dir, f"step_{step_idx + 1}_{action_timestamp}.png"),
                      "wb") as _f:
                _f.write(obs['screenshot'])

            with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                f.write(json.dumps({
                    "step_num": step_idx + 1,
                    "action_timestamp": action_timestamp,
                    "action": action,
                    "response": response,
                    "reward": reward,
                    "done": done,
                    "info": info,
                    "screenshot_file": f"step_{step_idx + 1}_{action_timestamp}.png"
                }))
                f.write("\n")
            if done:
                logger.info("The episode is done.")
                break
        step_idx += 1

    result = env.evaluate()
    logger.info("Result: %.2f", result)
    scores.append(result)
    with open(os.path.join(example_result_dir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{result}\n")
    env.controller.end_recording(os.path.join(example_result_dir, "recording.mp4"))

def run_single_example_autoglm(agent, env, example, max_steps, instruction, args, example_result_dir, scores):
    runtime_logger = setup_logger(example, example_result_dir)
    try:
        agent.reset(runtime_logger)
    except Exception as e:
        agent.reset()

    env.reset(task_config=example)
    
    time.sleep(60) # Wait for the environment to be ready
    obs = env._get_obs() # Get the initial observation
    done = False
    step_idx = 0
    env.controller.start_recording()
    while not done and step_idx < max_steps:
        response, actions = agent.predict(
            instruction,
            obs
        )
        for action in actions:
            # Capture the timestamp before executing the action
            action_timestamp = datetime.datetime.now().strftime("%Y%m%d@%H%M%S")
            logger.info("Step %d: %s", step_idx + 1, action)
            obs, reward, done, info = env.step(action, args.sleep_after_execution)

            logger.info("Reward: %.2f", reward)
            logger.info("Done: %s", done)
            # Save screenshot and trajectory information
            with open(os.path.join(example_result_dir, f"step_{step_idx + 1}_{action_timestamp}.png"),
                      "wb") as _f:
                _f.write(obs['screenshot'])
            with open(os.path.join(example_result_dir, "traj.jsonl"), "a") as f:
                f.write(json.dumps({
                    "step_num": step_idx + 1,
                    "action_timestamp": action_timestamp,
                    "action": action,
                    "response": response,
                    "reward": reward,
                    "done": done,
                    "info": info,
                    "screenshot_file": f"step_{step_idx + 1}_{action_timestamp}.png"
                }))
                f.write("\n")
            if done:
                logger.info("The episode is done.")
                break
        
        if not done: # not completed the task yet
            env.action_history.append('FAIL')
            
        step_idx += 1
    result = env.evaluate()
    logger.info("Result: %.2f", result)
    scores.append(result)
    with open(os.path.join(example_result_dir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{result}\n")
    env.controller.end_recording(os.path.join(example_result_dir, "recording.mp4"))
