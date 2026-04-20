#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 TabExplorer TE 图标
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_te_icon(filename='TabExplorer.ico', size=256):
    """
    创建一个带有 TE 字母的图标
    以 18px 为基准，按这个比例执行所有尺寸
    """
    from PIL import Image, ImageDraw, ImageFont
    
    blue = (33, 150, 243)  # #2196F3
    white = (255, 255, 255)

    # 蓝色方形背景（无圆角）
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(0, 0), (size - 1, size - 1)], fill=blue)

    # 内层白色区域（边框占 11%）
    margin = max(2, size * 11 // 100)
    inner_radius = 1
    draw.rounded_rectangle(
        [(margin, margin), (size - margin - 1, size - margin - 1)],
        radius=inner_radius,
        fill=white
    )

    # 蓝色 TE 文字
    text = "TE"
    font_size = max(5, int(size * 0.65))
    font = None
    for font_path in ["arialbd.ttf", "arial.ttf",
                      "C:\\Windows\\Fonts\\arialbd.ttf",
                      "C:\\Windows\\Fonts\\arial.ttf"]:
        try:
            font = ImageFont.truetype(font_path, font_size)
            break
        except:
            continue
    if font is None:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) // 2
    y = (size - text_height) // 2

    draw.text((x, y), text, font=font, fill=blue)
    
    # 保存为 ICO 文件
    icon_sizes = [(256, 256), (128, 128), (96, 96), (64, 64), (48, 48), (32, 32), (24, 24), (18, 18), (16, 16)]
    
    images = []
    for icon_size in icon_sizes:
        if icon_size == (256, 256):
            icons_img = img
        else:
            icons_img = img.resize(icon_size, Image.Resampling.LANCZOS)
        # 转为 RGB（ICO 格式需要）
        bg = Image.new('RGB', icons_img.size, (255, 255, 255))
        if icons_img.mode == 'RGBA':
            bg.paste(icons_img, mask=icons_img.split()[3])
        else:
            bg.paste(icons_img)
        images.append(bg)
    
    # 保存 ICO 文件
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
