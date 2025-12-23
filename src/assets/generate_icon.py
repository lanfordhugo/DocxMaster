#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标"""

from PIL import Image, ImageDraw, ImageFont
import os

def generate_icon():
    """生成一个简单的文档图标"""
    # 创建 256x256 的图像
    size = 256
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 绘制文档形状（带折角）
    doc_color = (52, 152, 219)  # 蓝色 #3498db
    margin = 40
    fold_size = 50
    
    # 文档主体多边形点
    points = [
        (margin, margin),  # 左上
        (size - margin - fold_size, margin),  # 右上（折角前）
        (size - margin, margin + fold_size),  # 折角点
        (size - margin, size - margin),  # 右下
        (margin, size - margin),  # 左下
    ]
    draw.polygon(points, fill=doc_color)
    
    # 绘制折角三角形（深色）
    fold_color = (41, 128, 185)  # 深蓝 #2980b9
    fold_points = [
        (size - margin - fold_size, margin),
        (size - margin, margin + fold_size),
        (size - margin - fold_size, margin + fold_size),
    ]
    draw.polygon(fold_points, fill=fold_color)
    
    # 绘制文本行（白色横线表示文字）
    line_color = (255, 255, 255, 200)
    line_y_start = 100
    line_spacing = 30
    line_margin = 70
    
    for i in range(4):
        y = line_y_start + i * line_spacing
        # 最后一行短一些
        end_x = size - line_margin - (40 if i == 3 else 0)
        draw.rectangle([line_margin, y, end_x, y + 8], fill=line_color)
    
    # 保存为 ICO 和 PNG
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 保存多尺寸 ICO
    ico_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    ico_images = [img.resize(s, Image.Resampling.LANCZOS) for s in ico_sizes]
    ico_path = os.path.join(script_dir, 'app_icon.ico')
    ico_images[0].save(ico_path, format='ICO', sizes=ico_sizes, append_images=ico_images[1:])
    print(f"已生成: {ico_path}")
    
    # 保存 PNG
    png_path = os.path.join(script_dir, 'app_icon.png')
    img.save(png_path, format='PNG')
    print(f"已生成: {png_path}")

if __name__ == '__main__':
    generate_icon()
