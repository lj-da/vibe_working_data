#!/bin/bash

# OSWorld 人工操作模式示例脚本
# Manual Operation Mode Example Script for OSWorld

echo "======================================================================"
echo "🚀 OSWorld 人工操作模式 / OSWorld Manual Operation Mode"
echo "======================================================================"
echo ""
echo "📋 此脚本将启动OSWorld环境，允许您手动执行任务"
echo "📋 This script will start OSWorld environment for manual task execution"
echo ""
echo "⚙️  配置说明 / Configuration:"
echo "   - 只执行前3个任务 / Execute only first 3 tasks"
echo "   - 使用Docker环境 / Using Docker environment"
echo "   - 启用GUI显示 / GUI display enabled"
echo "   - 人工操作模式 / Manual operation mode"
echo ""
echo "💡 使用提示 / Usage Tips:"
echo "   1. 脚本启动后会显示任务描述"
echo "   2. 在虚拟机窗口中手动完成任务"
echo "   3. 完成后返回命令行按回车键"
echo "   4. 系统将自动评估任务完成情况"
echo ""
echo "   1. Task description will be shown after startup"
echo "   2. Manually complete the task in VM window"
echo "   3. Return to command line and press Enter when done"
echo "   4. System will automatically evaluate task completion"
echo ""

# 等待用户确认
read -p "按回车键开始 / Press Enter to start..." -r

echo ""
echo "🔧 启动OSWorld环境... / Starting OSWorld environment..."
echo ""

# 运行OSWorld人工操作模式
python3 run_multienv_manual.py \
    --provider_name docker \
    --enable_gui \
    --headless false \
    --max_tasks 3 \
    --num_envs 1 \
    --action_space pyautogui \
    --observation_type screenshot \
    --enable_network \
    --model manual_operation \
    --result_dir ./results_manual \
    --domain chrome \
    --log_level INFO

echo ""
echo "✅ 任务执行完成！/ Task execution completed!"
echo "📊 结果保存在 ./results_manual 目录中"
echo "📊 Results saved in ./results_manual directory"
echo "======================================================================"

