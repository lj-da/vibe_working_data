#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
在图片上绘制红点的脚本
读取image_3.png，在归一化坐标(0-1000) [980.87, 81.099]位置画红点，保存为image_3_label.png
"""

import cv2
import numpy as np
import os

def draw_point_on_image(image_path, output_path, normalized_x, normalized_y, point_size=5, color=(0, 0, 255)):
    """
    在图片上绘制红点
    
    Args:
        image_path: 输入图片路径
        output_path: 输出图片路径
        normalized_x: 归一化x坐标 (0-1000)
        normalized_y: 归一化y坐标 (0-1000)
        point_size: 点的大小
        color: 点的颜色 (BGR格式)
    """
    # 检查输入图片是否存在
    if not os.path.exists(image_path):
        print(f"错误: 图片文件 {image_path} 不存在!")
        return False
    
    # 读取图片
    image = cv2.imread(image_path)
    if image is None:
        print(f"错误: 无法读取图片 {image_path}")
        return False
    
    # 获取图片尺寸
    height, width = image.shape[:2]
    print(f"图片尺寸: {width} x {height}")
    
    # 将归一化坐标(0-1000)转换为实际像素坐标
    actual_x = int((normalized_x / 1000.0) * width)
    actual_y = int((normalized_y / 1000.0) * height)
    
    print(f"归一化坐标: ({normalized_x}, {normalized_y})")
    print(f"实际像素坐标: ({actual_x}, {actual_y})")
    
    # 确保坐标在图片范围内
    actual_x = max(0, min(actual_x, width - 1))
    actual_y = max(0, min(actual_y, height - 1))
    
    # 在图片上画红点
    cv2.circle(image, (actual_x, actual_y), point_size, color, -1)
    
    # 可选：画一个更大的空心圆来突出显示
    cv2.circle(image, (actual_x, actual_y), point_size + 2, color, 2)
    
    # 保存图片
    success = cv2.imwrite(output_path, image)
    if success:
        print(f"成功保存标注图片到: {output_path}")
        return True
    else:
        print(f"错误: 保存图片失败 {output_path}")
        return False

def main():
    # 设置文件路径
    input_image = "image_3.png"
    output_image = "image_3_label.png"
    
    # 归一化坐标 (0-1000范围)
    norm_x = 980.87
    norm_y = 81.099
    
    print("开始处理图片...")
    print(f"输入图片: {input_image}")
    print(f"输出图片: {output_image}")
    
    # 绘制红点
    success = draw_point_on_image(
        image_path=input_image,
        output_path=output_image,
        normalized_x=norm_x,
        normalized_y=norm_y,
        point_size=5,
        color=(0, 0, 255)  # 红色 (BGR格式)
    )
    
    if success:
        print("✅ 脚本执行完成!")
    else:
        print("❌ 脚本执行失败!")

if __name__ == "__main__":
    main()
