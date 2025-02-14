import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pptx import Presentation as PptxPresentation
import os
import time
from pptx.util import Pt
import traceback
from copy import deepcopy
from PIL import Image, ImageDraw, ImageFont
import io
from pptx.enum.shapes import MSO_SHAPE_TYPE
import win32com.client
import pythoncom
from spire.presentation import Presentation
import comtypes.client

class PPTGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书图文批量制作工具")
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # PPT模板文件路径
        ttk.Label(self.main_frame, text="PPT模板文件路径:").grid(row=0, column=0, sticky=tk.W)
        self.ppt_path = tk.StringVar()
        self.ppt_entry = ttk.Entry(self.main_frame, textvariable=self.ppt_path, width=40)
        self.ppt_entry.grid(row=0, column=1, padx=5)
        ttk.Button(self.main_frame, text="选择", command=self.select_ppt).grid(row=0, column=2)
        
        # Excel文件路径
        ttk.Label(self.main_frame, text="Excel数据文件路径:").grid(row=1, column=0, sticky=tk.W)
        self.excel_path = tk.StringVar()
        self.excel_entry = ttk.Entry(self.main_frame, textvariable=self.excel_path, width=40)
        self.excel_entry.grid(row=1, column=1, padx=5)
        ttk.Button(self.main_frame, text="选择", command=self.select_excel).grid(row=1, column=2)
        
        # 保存路径
        ttk.Label(self.main_frame, text="保存文件路径:").grid(row=2, column=0, sticky=tk.W)
        self.save_path = tk.StringVar()
        self.save_entry = ttk.Entry(self.main_frame, textvariable=self.save_path, width=40)
        self.save_entry.grid(row=2, column=1, padx=5)
        ttk.Button(self.main_frame, text="选择", command=self.select_save_path).grid(row=2, column=2)
        
        # 单选按钮组
        self.radio_frame = ttk.LabelFrame(self.main_frame, text="特定页设置", padding="5")
        self.radio_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.radio_var1 = tk.StringVar(value="option1")
        self.radio_var2 = tk.StringVar(value="option1")
        
        ttk.Radiobutton(self.radio_frame, text="标题+正文", variable=self.radio_var1, value="option1").grid(row=0, column=0)
        ttk.Radiobutton(self.radio_frame, text="只有正文", variable=self.radio_var1, value="option2").grid(row=0, column=1)
        
        ttk.Radiobutton(self.radio_frame, text="所有页面一一对应", variable=self.radio_var2, value="option1").grid(row=1, column=0)
        ttk.Radiobutton(self.radio_frame, text="所有页面统一一个标题", variable=self.radio_var2, value="option2").grid(row=1, column=1)
        
        # 添加图片设置框
        self.scale_frame = ttk.LabelFrame(self.main_frame, text="图片设置", padding="5")
        self.scale_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 图片宽度设置
        ttk.Label(self.scale_frame, text="图片宽度:").grid(row=0, column=0, sticky=tk.W)
        self.width_var = tk.StringVar(value="2000")  # 默认宽度2000
        self.width_entry = ttk.Entry(self.scale_frame, textvariable=self.width_var, width=10)
        self.width_entry.grid(row=0, column=1, padx=5)
        
        # 图片高度设置
        ttk.Label(self.scale_frame, text="图片高度:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.height_var = tk.StringVar(value="1500")  # 默认高度1500
        self.height_entry = ttk.Entry(self.scale_frame, textvariable=self.height_var, width=10)
        self.height_entry.grid(row=0, column=3, padx=5)
        
        # 生成按钮
        ttk.Button(self.main_frame, text="开始批量生成", command=self.generate_ppt).grid(row=5, column=0, columnspan=3, pady=10)

    def select_ppt(self):
        filename = filedialog.askopenfilename(
            title="选择PPT模板文件",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if filename:
            self.ppt_path.set(filename)

    def select_excel(self):
        filename = filedialog.askopenfilename(
            title="选择Excel数据文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if filename:
            self.excel_path.set(filename)

    def select_save_path(self):
        dirname = filedialog.askdirectory(
            title="选择保存文件夹"
        )
        if dirname:
            # 生成默认文件名（当前时间）
            current_time = time.strftime("%Y%m%d_%H%M%S")
            default_filename = f"小红书图文_{current_time}.pptx"
            # 组合完整的保存路径
            save_path = os.path.join(dirname, default_filename)
            self.save_path.set(save_path)

    def print_shape_info(self, ppt):
        for slide in ppt.slides:
            print("\n=== 幻灯片信息 ===")
            for shape in slide.shapes:
                print(f"形状名称: {shape.name}")
                print(f"形状类型: {shape.shape_type}")
                if hasattr(shape, "text"):
                    print(f"文本内容: {shape.text}")
                if hasattr(shape, "placeholder_format"):
                    print(f"占位符类型: {shape.placeholder_format.type}")
                print("---")

    def convert_ppt_to_images(self, ppt_path, batch_size=50):
        try:
            # 获取文件名（不含扩展名）作为文件夹名
            base_name = os.path.splitext(os.path.basename(ppt_path))[0]
            
            # 创建与PPT同名的文件夹
            ppt_dir = os.path.dirname(ppt_path)
            images_dir = os.path.join(ppt_dir, base_name)
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)

            # 初始化COM
            pythoncom.CoInitialize()
            
            try:
                # 启动 WPS 应用程序
                wps = comtypes.client.CreateObject("KWPP.Application")
                wps.Visible = True  # 设置为可见，避免一些COM错误
                
                # 打开PPT文件
                ppt = wps.Presentations
                presentation = ppt.Open(os.path.abspath(ppt_path))  # 使用绝对路径
                
                slide_count = presentation.Slides.Count
                for batch_start in range(0, slide_count, batch_size):
                    for i in range(batch_start, min(batch_start + batch_size, slide_count)):
                        slide = presentation.Slides.Item(i + 1)  # 使用Item方法
                        # 生成输出文件路径
                        output_path = os.path.join(images_dir, f"{base_name}_第{i+1}页.jpg")
                        
                        # 导出当前幻灯片为JPG格式
                        slide.Export(output_path, "JPG")
                        print(f"已将幻灯片 {i + 1} 保存为 {output_path}")
                        
                        # 使用Pillow来提高清晰度（DPI）
                        img = Image.open(output_path)
                        img.save(output_path, "JPEG", quality=95, dpi=(300, 300))
                    
                    print(f"已处理第{batch_start + 1}到{min(batch_start + batch_size, slide_count)}张幻灯片")
                
                return images_dir
                
            finally:
                # 清理资源
                try:
                    if 'presentation' in locals():
                        presentation.Close()
                    if 'wps' in locals():
                        wps.Quit()
                except:
                    pass
                pythoncom.CoUninitialize()
                
        except Exception as e:
            raise Exception(f"转换图片时出错: {str(e)}")

    def generate_ppt(self):
        try:
            # 验证文件路径
            if not self.ppt_path.get():
                messagebox.showerror("错误", "请选择PPT模板文件")
                return
            if not self.excel_path.get():
                messagebox.showerror("错误", "请选择Excel数据文件")
                return
            if not self.save_path.get():
                messagebox.showerror("错误", "请选择保存位置")
                return
            
            # 验证文件是否存在
            if not os.path.exists(self.ppt_path.get()):
                messagebox.showerror("错误", "PPT模板文件不存在")
                return
            if not os.path.exists(self.excel_path.get()):
                messagebox.showerror("错误", "Excel数据文件不存在")
                return
            
            # 验证保存路径的文件夹是否存在
            save_dir = os.path.dirname(self.save_path.get())
            if not os.path.exists(save_dir):
                messagebox.showerror("错误", "保存文件夹不存在")
                return

            # 读取Excel文件
            df = pd.read_excel(self.excel_path.get())
            
            # 使用python-pptx的Presentation
            ppt = PptxPresentation(self.ppt_path.get())
            
            # 获取单选按钮的值
            has_title = self.radio_var1.get() == "option1"  # 是否包含标题
            unified_title = self.radio_var2.get() == "option2"  # 是否统一标题
            
            # 获取模板第一页
            template_slide = ppt.slides[0]
            
            # 获取Excel的第一行（用于匹配PPT中的对象）
            headers = df.iloc[0]
            
            # 遍历Excel数据（从第二行开始）
            for index, row in df.iloc[1:].iterrows():
                # 复制整个幻灯片
                new_slide = ppt.slides.add_slide(template_slide.slide_layout)
                
                # 删除新幻灯片中的默认形状
                for shape in new_slide.shapes:
                    element = shape._element
                    element.getparent().remove(element)
                
                # 从模板导入所有形状（包括格式和背景）
                for shape in template_slide.shapes:
                    el = shape.element
                    new_el = deepcopy(el)
                    new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
                
                # 遍历所有形状并只更新文本内容
                for shape in new_slide.shapes:
                    if not hasattr(shape, 'text_frame'):
                        continue
                    
                    # 获取shape的名称
                    shape_name = shape.name
                    print(f"当前形状名称: {shape_name}")  # 调试信息
                    
                    # 在Excel的列名中查找匹配的内容
                    for col, header_text in headers.items():
                        if str(col).strip() == shape_name.strip():
                            # 找到匹配的列，获取对应的内容
                            content = str(row[col])
                            print(f"匹配到列: {col}, 内容: {content}")  # 调试信息
                            
                            # 如果是标题且选择了"只有正文"，则跳过
                            if "标题" in shape_name and not has_title:
                                continue
                                
                            # 如果是标题且选择了"统一标题"
                            if "标题" in shape_name and unified_title:
                                content = str(df.iloc[1][col]) if index == 1 else shape.text_frame.text
                            
                            # 只更新文本内容，保持原有格式
                            for paragraph in shape.text_frame.paragraphs:
                                if paragraph.runs:
                                    # 保持原有格式，只更新文本
                                    paragraph.runs[0].text = content
                                else:
                                    # 如果没有runs，创建新的并复制原有格式
                                    run = paragraph.add_run()
                                    run.text = content
                            break
            
            # 保存生成的PPT
            ppt.save(self.save_path.get())
            
            # 转换为图片
            images_dir = self.convert_ppt_to_images(self.save_path.get())
            
            # 自动打开生成的PPT文件和图片文件夹
            try:
                os.startfile(self.save_path.get())
                os.startfile(images_dir)
            except Exception as open_error:
                print(f"打开文件失败: {str(open_error)}")
            
            messagebox.showinfo("成功", "PPT生成完成！图片已保存到images文件夹")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成过程中出现错误：{str(e)}\n{traceback.format_exc()}")

def main():
    root = tk.Tk()
    app = PPTGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 