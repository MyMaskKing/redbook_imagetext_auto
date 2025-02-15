from PIL import Image, ImageDraw, ImageFont
import os

def create_redbook_icon():
    # 创建一个512x512的图像
    size = 512
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制渐变背景
    margin = size * 0.1
    circle_bbox = (margin, margin, size - margin, size - margin)
    
    # 创建科技感的渐变背景
    for i in range(int(margin), int(size - margin)):
        alpha = (i - margin) / (size - 2 * margin)
        color = (
            int(255 * (1 - alpha) + 255 * alpha),  # R: 255 -> 255
            int(36 * (1 - alpha) + 66 * alpha),    # G: 36 -> 66
            int(66 * (1 - alpha) + 97 * alpha)     # B: 66 -> 97
        )
        draw.ellipse(
            (margin, i, size - margin, i + 1),
            fill=color
        )
    
    # 添加科技感装饰
    # 绘制四个角的装饰线
    line_length = size * 0.15
    line_width = 3
    corner_margin = size * 0.2
    
    # 左上角
    draw.line([(corner_margin, corner_margin), (corner_margin + line_length, corner_margin)], fill='white', width=line_width)
    draw.line([(corner_margin, corner_margin), (corner_margin, corner_margin + line_length)], fill='white', width=line_width)
    
    # 右上角
    draw.line([(size - corner_margin - line_length, corner_margin), (size - corner_margin, corner_margin)], fill='white', width=line_width)
    draw.line([(size - corner_margin, corner_margin), (size - corner_margin, corner_margin + line_length)], fill='white', width=line_width)
    
    # 左下角
    draw.line([(corner_margin, size - corner_margin), (corner_margin + line_length, size - corner_margin)], fill='white', width=line_width)
    draw.line([(corner_margin, size - corner_margin - line_length), (corner_margin, size - corner_margin)], fill='white', width=line_width)
    
    # 右下角
    draw.line([(size - corner_margin - line_length, size - corner_margin), (size - corner_margin, size - corner_margin)], fill='white', width=line_width)
    draw.line([(size - corner_margin, size - corner_margin - line_length), (size - corner_margin, size - corner_margin)], fill='white', width=line_width)
    
    try:
        # 尝试加载微软雅黑字体
        title_font = ImageFont.truetype("msyh.ttc", int(size * 0.12))  # 减小字体大小
    except:
        title_font = ImageFont.load_default()
    
    # 添加文字
    text = "小红书图文\n批量制作工具"  # 分两行显示
    
    # 获取文字大小
    text_bbox = draw.textbbox((0, 0), text, font=title_font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    
    # 计算文字位置使其居中
    x = (size - text_width) / 2
    y = (size - text_height) / 2
    
    # 绘制文字阴影
    shadow_offset = 2
    draw.text((x + shadow_offset, y + shadow_offset), text, fill=(0, 0, 0, 100), font=title_font)
    
    # 绘制主文字
    draw.text((x, y), text, fill='white', font=title_font)
    
    # 保存为多种尺寸的ICO文件
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    icons = []
    for s in sizes:
        icons.append(image.resize(s, Image.Resampling.LANCZOS))
    
    # 保存ICO文件
    icons[0].save(
        'redbook.ico',
        format='ICO',
        sizes=sizes,
        append_images=icons[1:]
    )

if __name__ == '__main__':
    create_redbook_icon() 