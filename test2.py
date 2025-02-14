import os
import comtypes.client
from PIL import Image


def ppt_to_png(ppt_file, output_dir, batch_size=50):
    # 启动 WPS 应用程序
    wps = comtypes.client.CreateObject("KWPP.Application")  # 使用KWPP而不是kWPS
    wps.Visible = True  # 设置WPS应用为可见
    
    try:
        # 打开PPT文件
        ppt = wps.Presentations
        presentation = ppt.Open(ppt_file)
        
        # 设置输出文件夹
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        slide_count = presentation.Slides.Count
        for batch_start in range(0, slide_count, batch_size):
            for i in range(batch_start, min(batch_start + batch_size, slide_count)):
                slide = presentation.Slides.Item(i + 1)  # 使用Item方法
                # 生成临时文件路径
                temp_img_path = os.path.join(output_dir, f"slide_{i + 1}.png")
                
                # 导出当前幻灯片为PNG格式
                slide.Export(temp_img_path, "PNG")
                print(f"已将幻灯片 {i + 1} 保存为 {temp_img_path}")
                
                # 使用Pillow来提高清晰度（DPI）
                img = Image.open(temp_img_path)
                img.save(temp_img_path, dpi=(300, 300))  # 调整为300 DPI
            
            print(f"已处理第{batch_start + 1}到{min(batch_start + batch_size, slide_count)}张幻灯片\n")
        
    finally:
        # 关闭PPT文件和WPS
        try:
            presentation.Close()
            wps.Quit()
        except:
            pass


# 定义PPT文件路径和目标路径
ppt_file_path = r"D:\AboutDev\Workspace_cursor\redbook_imagetext_auto\image\小红书图文_20250214_202931.pptx"
output_dir = r"D:\AboutDev\Workspace_cursor\redbook_imagetext_auto\Output"

# 调用函数转换PPT为PNG
ppt_to_png(ppt_file_path, output_dir)