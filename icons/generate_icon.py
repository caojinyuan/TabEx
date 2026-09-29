#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 TabExplorer TE 图标
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_te_icon(filename='TabExplorer.ico', size=256):
    """
    创建蓝白圆角主题的 TE 图标
    按 256px 基准等比例生成各尺寸
    """
    from PIL import Image, ImageDraw, ImageFont
    
    blue = (33, 150, 243)  # #2196F3
    white = (255, 255, 255)

    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 外层蓝色圆角背景
    outer_radius = max(2, size * 40 // 256)
    draw.rounded_rectangle(
        [(0, 0), (size - 1, size - 1)],
        radius=outer_radius,
        fill=blue
    )
    
    # 内层白色圆角容器（形成蓝色边框效果）
    margin = max(2, size * 28 // 256)
    inner_radius = max(2, size * 24 // 256)
    draw.rounded_rectangle(
        [(margin, margin), (size - margin - 1, size - margin - 1)],
        radius=inner_radius,
        fill=white
    )

    # 蓝色 TE 文字（精确填充内框高度，压窄宽度使其瘦高）
    text = "TE"
    font = None
    for font_path in ["arialbd.ttf", "arial.ttf",
                      "C:\\Windows\\Fonts\\arialbd.ttf",
                      "C:\\Windows\\Fonts\\arial.ttf"]:
        try:
            font = ImageFont.truetype(font_path, 512)
            break
        except:
            continue
    if font is None:
        font = ImageFont.load_default()

    # 在大画布上渲染，裁剪实际文字像素边界
    big = Image.new('RGBA', (1024, 1024), (0, 0, 0, 0))
    big_draw = ImageDraw.Draw(big)
    big_draw.text((256, 256), text, font=font, fill=blue)
    bbox = big.getbbox()
    if bbox:
        text_crop = big.crop(bbox)
        inner = size - 2 * max(2, size * 28 // 256)
        target_h = int(inner * 0.82)   # 高度：内框的 82%
        target_w = int(target_h * 0.90)  # 宽度：高度的 90%，使 TE 瘦高
        text_resized = text_crop.resize((target_w, target_h), Image.Resampling.LANCZOS)
        px = (size - target_w) // 2
        py = (size - target_h) // 2
        img.paste(text_resized, (px, py), text_resized)
    
    # 保存为 ICO 文件
    icon_sizes = [(256, 256), (128, 128), (96, 96), (64, 64), (48, 48), (32, 32), (24, 24), (18, 18), (16, 16)]
    
    images = []
    for icon_size in icon_sizes:
        if icon_size == (256, 256):
            icons_img = img
        else:
            icons_img = img.resize(icon_size, Image.Resampling.LANCZOS)
        # 保持 RGBA 模式以支持透明度
        images.append(icons_img)
    
    # 保存 ICO 文件（RGBA 格式支持透明）
    images[0].save(
        filename,
        format='ICO',
        sizes=[(img.width, img.height) for img in images]
    )
    
    print(f"✓ 图标已生成: {filename}")
    print(f"  图标大小: {size}x{size} 像素")
    print(f"  样式: 白底、蓝色方形边框、蓝色 TE 文字")
    print(f"  支持的分辨率: {', '.join([f'{s[0]}x{s[1]}' for s in icon_sizes])}")
    return filename

if __name__ == '__main__':
    # 生成图标
    icon_path = os.path.join(os.path.dirname(__file__), 'TabExplorer.ico')
    create_te_icon(icon_path)
    
    if os.path.exists(icon_path):
        file_size = os.path.getsize(icon_path)
        print(f"  文件大小: {file_size} 字节")
        print("\n现在可以运行 2_build_exe.bat 来打包应用，图标将被应用到 exe 文件！")
    else:
        print("错误：图标生成失败")
