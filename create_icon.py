from PIL import Image, ImageDraw
import os, math

def create_redbook_icon():
    # 基础设置
    size = 1024
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制圆形背景
    margin = size * 0.05
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill='#FF2442'  # 小红书品牌红色
    )
    
    # 计算白色方块的尺寸和位置
    square_margin = size * 0.25
    square_size = size - 2 * square_margin
    
    # 绘制白色方块
    draw.rounded_rectangle(
        [square_margin, square_margin, 
         size - square_margin, size - square_margin],
        radius=square_size * 0.15,
        fill='white'
    )
    
    # 计算中心位置
    center_x = size / 2
    center_y = size / 2
    
    # 绘制渐变装饰线条
    line_count = 8
    line_spacing = square_size / (line_count + 1)
    
    for direction in [1, -1]:
        for i in range(line_count):
            start_x = square_margin + line_spacing * (i + 0.5)
            if direction == -1:
                start_x = size - start_x
            
            segments = 25
            for j in range(segments):
                progress = j / segments
                offset = (math.sin(progress * math.pi) * 0.3 + 
                         math.sin(progress * math.pi * 2) * 0.1) * line_spacing
                
                y1 = square_margin + square_size * progress * 0.7
                y2 = y1 + square_size * 0.7 / segments
                
                alpha = int(35 * (1 - progress * 0.6))
                x = start_x + offset * direction
                
                draw.line(
                    [x, y1, x, y2],
                    fill=(255, 36, 66, alpha),
                    width=int(size * 0.002)
                )
    
    # 绘制装饰点
    for i in range(12):
        angle = i * math.pi / 6
        for radius_factor in [0.32, 0.38]:
            radius = square_size * radius_factor
            x = center_x + math.cos(angle) * radius
            y = center_y + math.sin(angle) * radius
            dot_size = size * 0.006
            
            for j in range(4):
                alpha = int(35 - j * 8)
                current_size = dot_size * (1 + j * 0.4)
                draw.ellipse(
                    [x - current_size, y - current_size,
                     x + current_size, y + current_size],
                    fill=(255, 36, 66, alpha)
                )
    
    # 绘制中心红色方块
    inner_size = square_size * 0.35
    draw.rounded_rectangle(
        [center_x - inner_size/2, center_y - inner_size/2,
         center_x + inner_size/2, center_y + inner_size/2],
        radius=inner_size * 0.2,
        fill='#FF2442'
    )
    
    # 机器人设计
    robot_size = inner_size * 0.8
    head_width = robot_size * 0.7
    head_height = robot_size * 0.6
    
    # 计算机器人位置（确保所有坐标都是正值）
    head_y = center_y - head_height/2  # 将头部居中
    
    # 添加头部光晕效果
    for i in range(4):
        offset = i * 1.5
        alpha = int(50 - i * 12)
        draw.rounded_rectangle(
            [center_x - head_width/2 - offset, head_y - offset,
             center_x + head_width/2 + offset, head_y + head_height + offset],
            radius=(head_width * 0.4 + offset),
            fill=(255, 255, 255, alpha)
        )
    
    # 主头部
    draw.rounded_rectangle(
        [center_x - head_width/2, head_y,
         center_x + head_width/2, head_y + head_height],
        radius=head_width * 0.4,
        fill='white'
    )
    
    # 天线
    antenna_width = head_width * 0.08
    antenna_height = head_height * 0.25
    antenna_spacing = head_width * 0.25
    
    for x_offset in [-antenna_spacing, antenna_spacing]:
        antenna_x = center_x + x_offset
        # 天线光晕
        for i in range(3):
            offset = i * 1
            alpha = int(50 - i * 15)
            draw.rounded_rectangle(
                [antenna_x - antenna_width/2 - offset, head_y - antenna_height - offset,
                 antenna_x + antenna_width/2 + offset, head_y + offset],
                radius=(antenna_width/2 + offset),
                fill=(255, 255, 255, alpha)
            )
        # 主天线
        draw.rounded_rectangle(
            [antenna_x - antenna_width/2, head_y - antenna_height,
             antenna_x + antenna_width/2, head_y],
            radius=antenna_width/2,
            fill='white'
        )
    
    # 眼睛
    eye_width = head_width * 0.25
    eye_height = head_height * 0.15
    eye_spacing = head_width * 0.2
    eye_y = head_y + head_height * 0.3
    
    for x_offset in [-eye_spacing, eye_spacing]:
        eye_x = center_x + x_offset
        # 眼睛光晕
        for i in range(3):
            offset = i * 1
            alpha = int(80 - i * 25)
            draw.rounded_rectangle(
                [eye_x - eye_width/2 - offset, eye_y - offset,
                 eye_x + eye_width/2 + offset, eye_y + eye_height + offset],
                radius=(eye_height/2 + offset),
                fill=(255, 36, 66, alpha)
            )
        # 主眼睛
        draw.rounded_rectangle(
            [eye_x - eye_width/2, eye_y,
             eye_x + eye_width/2, eye_y + eye_height],
            radius=eye_height/2,
            fill='#FF2442'
        )
    
    # 保存图标
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    icons = []
    
    for s in sizes:
        temp_size = (s[0] * 2, s[1] * 2)
        temp_image = image.resize(temp_size, Image.Resampling.LANCZOS)
        icons.append(temp_image.resize(s, Image.Resampling.LANCZOS))
    
    icons[0].save(
        'redbook.ico',
        format='ICO',
        sizes=sizes,
        append_images=icons[1:]
    )

if __name__ == '__main__':
    create_redbook_icon() 