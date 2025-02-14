from PIL import Image, ImageDraw, ImageFont
import os

def create_redbook_icon():
    # 创建一个512x512的图像（推荐的图标尺寸）
    size = 512
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制圆形背景
    margin = size * 0.1
    circle_bbox = (margin, margin, size - margin, size - margin)
    draw.ellipse(circle_bbox, fill='#FF2442')
    
    # 添加文字
    try:
        # 尝试加载微软雅黑字体
        font = ImageFont.truetype("msyh.ttc", int(size * 0.4))
    except:
        # 如果找不到，使用默认字体
        font = ImageFont.load_default()
    
    text = "图文"
    # 获取文字大小
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    
    # 计算文字位置使其居中
    x = (size - text_width) / 2
    y = (size - text_height) / 2
    
    # 绘制文字
    draw.text((x, y), text, fill='white', font=font)
    
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