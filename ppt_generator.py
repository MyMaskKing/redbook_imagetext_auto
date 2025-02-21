import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import time
import traceback
from copy import deepcopy
from PIL import Image
import comtypes.client
import requests
import win32gui
import win32con

class ModernButton(tk.Button):
    def __init__(self, master, **kwargs):
        # 提取自定义颜色参数
        self.start_color = kwargs.pop('start_color', '#FF4D6D') if 'start_color' in kwargs else '#FF4D6D'
        self.end_color = kwargs.pop('end_color', '#FF8FA3') if 'end_color' in kwargs else '#FF8FA3'
        
        # 设置基本配置
        kwargs.update({
            'background': self.start_color,
            'foreground': 'white',
            'font': ('Microsoft YaHei UI', 10),
            'borderwidth': 0,
            'activebackground': self.end_color,
            'activeforeground': 'white',
            'padx': 15,
            'pady': 8,
            'cursor': 'hand2',
            'relief': 'flat'
        })
        
        super().__init__(master, **kwargs)
        
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)

    def on_enter(self, e):
        self.config(background=self.end_color)

    def on_leave(self, e):
        self.config(background=self.start_color)

class PPTGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书图文批量制作工具")
        
        # 设置窗口大小和位置
        window_width = 800
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # 设置小红书风格主题（浅粉色背景）
        self.root.configure(bg='#FFF0F5')
        
        # 创建主容器
        main_container = tk.Frame(root, bg='#FFF0F5')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(main_container, bg='#FFF0F5', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        
        # 创建主框架
        self.main_frame = tk.Frame(self.canvas, bg='#FFF0F5')
        
        # 配置画布滚动区域
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # 绑定画布和滚动条
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 布局画布和滚动条
        self.canvas.pack(side="left", fill="both", expand=True, padx=(20, 0))  # 添加左边距
        self.scrollbar.pack(side="right", fill="y")
        
        # 绑定事件
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.main_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 标题（使用小红书logo颜色）
        title_label = tk.Label(
            self.main_frame,
            text="小红书图文批量制作工具",
            font=('Microsoft YaHei UI', 24, 'bold'),
            fg='#FF2442',
            bg='#FFF0F5'
        )
        title_label.pack(pady=(0, 30))
        
        # 文件选择区域
        self.create_file_frame()
        
        # 添加AI提问模板区域
        self.create_ai_template_frame()
        
        # 设置区域
        self.create_settings_frame()
        
        # 添加尺寸设置区域
        self.create_scale_frame()
        
        # 进度条区域
        self.create_progress_frame()
        
        # WPS提示
        self.create_wps_notice()
        
        # 生成按钮
        self.create_generate_button()

    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_frame_configure(self, event=None):
        """更新画布的滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """当画布大小改变时调整框架宽度"""
        # 设置框架宽度以匹配画布
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def create_file_frame(self):
        file_frame = tk.LabelFrame(
            self.main_frame,
            text="文件选择",
            font=('Microsoft YaHei UI', 12, 'bold'),
            fg='#FF2442',
            bg='#FFFFFF',
            padx=20,
            pady=20,
            relief='flat',
        )
        file_frame.pack(fill=tk.X, pady=(0, 20), padx=20)

        # 配置列的权重
        file_frame.grid_columnconfigure(1, weight=1)  # 让输入框列自动扩展

        # 为每个选择按钮设置不同的渐变色
        self.ppt_path = self.create_file_entry(
            file_frame, "PPT模板:", self.select_ppt, 0,
            start_color='#FF4D6D', end_color='#FF8FA3'
        )
        
        self.excel_path = self.create_file_entry(
            file_frame, "Excel文件:", self.select_excel, 1,
            start_color='#FF6B6B', end_color='#FFA5A5'
        )
        
        self.save_path = self.create_file_entry(
            file_frame, "保存位置:", self.select_save_path, 2,
            start_color='#FF8882', end_color='#FFACAC'
        )

    def create_file_entry(self, parent, label_text, command, row, start_color, end_color):
        # 标签
        label = tk.Label(
            parent,
            text=label_text,
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            width=10,  # 固定标签宽度
            anchor='e'  # 右对齐
        )
        label.grid(row=row, column=0, sticky='e', pady=10, padx=(0, 10))
        
        # 输入框
        var = tk.StringVar()
        entry = tk.Entry(
            parent,
            textvariable=var,
            font=('Microsoft YaHei UI', 10),
            bg='#F8F8F8',
            fg='#333333',
            insertbackground='#666666',
            relief='flat',
            highlightthickness=1,
            highlightcolor='#FF4D6D',
            highlightbackground='#E0E0E0'
        )
        entry.grid(row=row, column=1, sticky='ew', padx=10)
        
        # 按钮
        button = ModernButton(
            parent,
            text="选择",
            command=command,
            start_color=start_color,
            end_color=end_color,
            width=8  # 固定按钮宽度
        )
        button.grid(row=row, column=2, padx=(0, 10))
        
        return var

    def create_settings_frame(self):
        settings_frame = tk.LabelFrame(
            self.main_frame,
            text="设置",
            font=('Microsoft YaHei UI', 12, 'bold'),
            fg='#FF2442',
            bg='#FFFFFF',
            padx=20,
            pady=20,
            relief='flat'
        )
        settings_frame.pack(fill=tk.X, pady=(0, 20), padx=20)

        # 使用Grid布局
        settings_frame.grid_columnconfigure(1, weight=1)
        
        # 标题设置
        tk.Label(
            settings_frame,
            text="标题设置：",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            anchor='w'
        ).grid(row=0, column=0, sticky='w', padx=(0, 10), pady=(0, 10))

        # 第一组单选按钮
        radio_frame1 = tk.Frame(settings_frame, bg='#FFFFFF')
        radio_frame1.grid(row=0, column=1, sticky='w')
        
        self.radio_var1 = tk.StringVar(value="option1")
        tk.Radiobutton(
            radio_frame1,
            text="包含标题",
            variable=self.radio_var1,
            value="option1",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            activebackground='#FFE4E8',
            selectcolor='#FF4D6D'
        ).pack(side=tk.LEFT, padx=(0, 20))

        tk.Radiobutton(
            radio_frame1,
            text="只有正文",
            variable=self.radio_var1,
            value="option2",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            activebackground='#FFE4E8',
            selectcolor='#FF4D6D'
        ).pack(side=tk.LEFT)

        # 标题处理
        tk.Label(
            settings_frame,
            text="标题处理：",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            anchor='w'
        ).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(0, 10))

        # 第二组单选按钮
        radio_frame2 = tk.Frame(settings_frame, bg='#FFFFFF')
        radio_frame2.grid(row=1, column=1, sticky='w')
        
        self.radio_var2 = tk.StringVar(value="option1")
        tk.Radiobutton(
            radio_frame2,
            text="每页不同",
            variable=self.radio_var2,
            value="option1",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            activebackground='#FFE4E8',
            selectcolor='#FF4D6D'
        ).pack(side=tk.LEFT, padx=(0, 20))

        tk.Radiobutton(
            radio_frame2,
            text="统一标题",
            variable=self.radio_var2,
            value="option2",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            activebackground='#FFE4E8',
            selectcolor='#FF4D6D'
        ).pack(side=tk.LEFT)

        # 首图字体大小
        tk.Label(
            settings_frame,
            text="首图字体大小：",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF',
            anchor='w'
        ).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(0, 10))
        
        # 创建输入框容器（用于实现更好的边框效果）
        entry_container = tk.Frame(settings_frame, bg='#E0E0E0', padx=1, pady=1)
        entry_container.grid(row=2, column=1, sticky='w')
        
        # 输入框
        self.font_size_var = tk.StringVar(value="45")  # 默认值为45
        self.font_size_entry = tk.Entry(
            entry_container,
            textvariable=self.font_size_var,
            font=('Microsoft YaHei UI', 10),
            width=5,
            relief='flat',  # 移除输入框自身的边框
            justify='center',  # 文字居中显示
            bg='#FFFFFF'
        )
        self.font_size_entry.pack(padx=1, pady=1)
        
        # 添加提示说明
        tip_label = tk.Label(
            settings_frame,
            text='提示：内容中包含"#我的首图#"的文本会自动调整为上方设置的字体大小',
            font=('Microsoft YaHei UI', 9),
            fg='#666666',
            bg='#FFFFFF'
        )
        tip_label.grid(row=3, column=0, columnspan=2, sticky='w', pady=(10, 0))

    def create_scale_frame(self):
        scale_frame = tk.LabelFrame(
            self.main_frame,
            text="图片尺寸设置",
            font=('Microsoft YaHei UI', 12, 'bold'),
            fg='#FF2442',
            bg='#FFFFFF',
            padx=25,
            pady=25,
            relief='flat'
        )
        scale_frame.pack(fill=tk.X, pady=(0, 20), padx=20)

        # 使用容器来居中对齐内容
        container = tk.Frame(scale_frame, bg='#FFFFFF')
        container.pack(expand=True, fill=tk.X, padx=15)
        
        # 预设尺寸选项
        size_options = [
            "自定义尺寸",
            "小红书封面（竖版）- 1080×1440",
            "小红书封面（横版）- 1440×1080",
            "小红书封面（方版）- 1080×1080",
            "小红书图文封面（竖版）- 1242×1660",
            "小红书图文封面（方版）- 1080×1080",
            "小红书图文封面（横版）- 2560×1440",
            "抖音视频封面 - 1080×1920",
            "抖音预览封面 - 1080×1464",
            "抖音个人主页背景 - 1125×633"
        ]
        
        # 尺寸映射字典
        self.size_mapping = {
            "自定义尺寸": None,
            "小红书封面（竖版）- 1080×1440": (1080, 1440),
            "小红书封面（横版）- 1440×1080": (1440, 1080),
            "小红书封面（方版）- 1080×1080": (1080, 1080),
            "小红书图文封面（竖版）- 1242×1660": (1242, 1660),
            "小红书图文封面（方版）- 1080×1080": (1080, 1080),
            "小红书图文封面（横版）- 2560×1440": (2560, 1440),
            "抖音视频封面 - 1080×1920": (1080, 1920),
            "抖音预览封面 - 1080×1464": (1080, 1464),
            "抖音个人主页背景 - 1125×633": (1125, 633)
        }
        
        # 创建居中容器
        center_frame = tk.Frame(container, bg='#FFFFFF')
        center_frame.pack(anchor='center', pady=(0, 20))
        
        # 预设尺寸组
        preset_frame = tk.Frame(center_frame, bg='#FFFFFF')
        preset_frame.pack(side=tk.LEFT)
        
        tk.Label(
            preset_frame,
            text="预设尺寸:",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF'
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        # 创建圆角下拉框容器
        combo_container = tk.Frame(preset_frame, bg='#E0E0E0', padx=1, pady=1)
        combo_container.pack(side=tk.LEFT)
        
        # 设置下拉框样式
        style = ttk.Style()
        style.configure(
            'Rounded.TCombobox',
            background='#FFFFFF',
            fieldbackground='#FFFFFF',
            foreground='#333333',
            arrowcolor='#FF2442',
            borderwidth=0,
            padding=8
        )
        
        self.size_var = tk.StringVar(value="小红书图文封面（竖版）- 1242×1660")
        size_combo = ttk.Combobox(
            combo_container,
            textvariable=self.size_var,
            values=size_options,
            state='readonly',
            width=25,
            font=('Microsoft YaHei UI', 10),
            style='Rounded.TCombobox'
        )
        size_combo.pack(padx=1, pady=1)
        
        # 创建第二行居中容器
        input_center_frame = tk.Frame(container, bg='#FFFFFF')
        input_center_frame.pack(anchor='center', pady=(10, 0))
        
        # 宽度输入组
        width_frame = tk.Frame(input_center_frame, bg='#FFFFFF')
        width_frame.pack(side=tk.LEFT, padx=(0, 30))
        
        tk.Label(
            width_frame,
            text="宽度:",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF'
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        # 创建圆角输入框容器
        width_entry_container = tk.Frame(width_frame, bg='#E0E0E0', padx=1, pady=1)
        width_entry_container.pack(side=tk.LEFT)
        
        self.width_var = tk.StringVar(value="1242")
        self.width_entry = tk.Entry(
            width_entry_container,
            textvariable=self.width_var,
            font=('Microsoft YaHei UI', 10),
            width=8,
            relief='flat',
            justify='center',
            bg='#FFFFFF'
        )
        self.width_entry.pack(padx=1, pady=1)
        
        # 高度输入组
        height_frame = tk.Frame(input_center_frame, bg='#FFFFFF')
        height_frame.pack(side=tk.LEFT)
        
        tk.Label(
            height_frame,
            text="高度:",
            font=('Microsoft YaHei UI', 10),
            fg='#333333',
            bg='#FFFFFF'
        ).pack(side=tk.LEFT, padx=(0, 8))

        # 添加链接到右上角
        link_label = tk.Label(
            scale_frame,
            text="小红书抖音图片尺寸说明",
            font=('Microsoft YaHei UI', 9, 'underline'),
            fg='#FF4D6D',
            bg='#FFFFFF',
            cursor='hand2'
        )
        link_label.place(relx=1.0, y=1.5, anchor='ne', x=-20)
        
        # 添加链接点击事件
        def open_link(event):
            import webbrowser
            webbrowser.open('https://kdocs.cn/l/cpCXrxJCZlzi?linkname=TSby1ZRlVS')
        
        link_label.bind('<Button-1>', open_link)
        
        # 创建圆角输入框容器
        height_entry_container = tk.Frame(height_frame, bg='#E0E0E0', padx=1, pady=1)
        height_entry_container.pack(side=tk.LEFT)
        
        self.height_var = tk.StringVar(value="1660")
        self.height_entry = tk.Entry(
            height_entry_container,
            textvariable=self.height_var,
            font=('Microsoft YaHei UI', 10),
            width=8,
            relief='flat',
            justify='center',
            bg='#FFFFFF'
        )
        self.height_entry.pack(padx=1, pady=1)
        
        # 添加输入框焦点事件
        def on_focus_in(event, container):
            container.configure(bg='#FF2442')
        
        def on_focus_out(event, container):
            container.configure(bg='#E0E0E0')
        
        self.width_entry.bind('<FocusIn>', lambda e: on_focus_in(e, width_entry_container))
        self.width_entry.bind('<FocusOut>', lambda e: on_focus_out(e, width_entry_container))
        self.height_entry.bind('<FocusIn>', lambda e: on_focus_in(e, height_entry_container))
        self.height_entry.bind('<FocusOut>', lambda e: on_focus_out(e, height_entry_container))
        
        # 绑定下拉框选择事件
        size_combo.bind('<<ComboboxSelected>>', self.on_size_selected)
        
        # 初始状态下禁用输入框
        self.width_entry.config(state='disabled')
        self.height_entry.config(state='disabled')

    def on_size_selected(self, event):
        """处理尺寸选择事件"""
        selected = self.size_var.get()
        if selected == "自定义尺寸":
            self.width_entry.config(state='normal')
            self.height_entry.config(state='normal')
        else:
            self.width_entry.config(state='disabled')
            self.height_entry.config(state='disabled')
            if selected in self.size_mapping:
                width, height = self.size_mapping[selected]
                self.width_var.set(str(width))
                self.height_var.set(str(height))

    def create_progress_frame(self):
        progress_frame = tk.Frame(self.main_frame, bg='#FFF0F5')
        progress_frame.pack(fill=tk.X, pady=(0, 20), padx=20)
        
        # 创建圆角进度条背景
        progress_bg = tk.Canvas(
            progress_frame,
            height=20,
            bg='#FFE4E8',
            highlightthickness=0
        )
        progress_bg.pack(fill=tk.X, padx=2)
        
        # 创建进度条
        self.progress_canvas = tk.Canvas(
            progress_bg,
            height=16,
            bg='#FFE4E8',
            highlightthickness=0
        )
        self.progress_canvas.place(relx=0.01, rely=0.5, relwidth=0.98, anchor='w')
        
        # 初始化进度变量
        self.progress_var = tk.DoubleVar(value=0)
        
        # 创建进度文本标签
        self.progress_label = tk.Label(
            progress_frame,
            text="准备就绪",
            font=('Microsoft YaHei UI', 10),
            fg='#666666',
            bg='#FFF0F5'
        )
        self.progress_label.pack(pady=10)

        # 绑定进度变量更新事件
        self.progress_var.trace_add('write', self._update_progress_bar)

    def _update_progress_bar(self, *args):
        """更新进度条显示"""
        progress = self.progress_var.get()
        width = self.progress_canvas.winfo_width()
        filled_width = int(width * (progress / 100))
        
        # 清除原有内容
        self.progress_canvas.delete('progress')
        
        # 绘制圆角进度条
        if filled_width > 0:
            # 计算圆角矩形的坐标
            x1, y1 = 0, 0
            x2, y2 = filled_width, 16
            radius = 8  # 圆角半径
            
            # 创建圆角矩形路径
            self.progress_canvas.create_polygon(
                x1+radius, y1,
                x2-radius, y1,
                x2, y1,
                x2, y2,
                x2-radius, y2,
                x1+radius, y2,
                x1, y2,
                x1, y1,
                fill='#FF4D6D',
                smooth=True,
                tags='progress'
            )

    def create_wps_notice(self):
        notice_frame = tk.Frame(self.main_frame, bg='#FFF0F5')
        notice_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(
            notice_frame,
            text="注意：本机必须安装WPS软件，否则无法生成图片",
            font=('Microsoft YaHei UI', 10),
            fg='#FF2442',
            bg='#FFE4E8',
            padx=15,
            pady=10
        ).pack(fill=tk.X)

    def create_ai_template_frame(self):
        """创建AI提问模板区域"""
        template_frame = tk.LabelFrame(
            self.main_frame,
            text="AI提问模板",
            font=('Microsoft YaHei UI', 12, 'bold'),
            fg='#FF2442',
            bg='#FFFFFF',
            padx=25,
            pady=25,
            relief='flat'
        )
        template_frame.pack(fill=tk.X, pady=(0, 20), padx=20)

        # 创建文本框容器，使用相对定位
        text_container = tk.Frame(template_frame, bg='#FFFFFF', padx=1, pady=1)
        text_container.pack(fill=tk.X)
        text_container.grid_propagate(False)  # 防止内容影响容器大小

        # 创建只读文本框
        template_text = tk.Text(
            text_container,
            font=('Microsoft YaHei UI', 10),
            bg='#F8F8F8',
            fg='#333333',
            height=4,
            relief='flat',
            padx=10,
            pady=10,
            wrap=tk.WORD
        )
        template_text.pack(fill=tk.X, expand=True)
        
        # 获取模板内容
        try:
            response = requests.get('https://webapi.mymaskking.us.kg/get_ai_template_hint', timeout=5)  # 添加5秒超时
            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 200:
                    template_content = data['data']['templates']
                else:
                    raise Exception('API返回状态错误')
            else:
                raise Exception('HTTP请求失败')
        except requests.Timeout:
            print("获取模板内容超时")
            template_content = """请帮我查找关于"今日的科技新闻"的内容，生成的格式为表格，有三列：标题，内容,并且帮我生成150字的小红书爆文，要求爆文标题和爆文内容足够吸引人眼球，里面可以插入一些表情"""
        except Exception as e:
            print(f"获取模板内容失败: {str(e)}")
            template_content = """请帮我查找关于"今日的科技新闻"的内容，生成的格式为表格，有三列：标题，内容,并且帮我生成150字的小红书爆文，要求爆文标题和爆文内容足够吸引人眼球，里面可以插入一些表情"""
        
        template_text.insert('1.0', template_content)
        template_text.config(state='disabled')

        # 创建复制按钮
        copy_button = tk.Label(
            text_container,
            text="复制",
            font=('Microsoft YaHei UI', 8),
            fg='#666666',
            bg='#EAEAEA',
            padx=10,
            pady=2,
            cursor='hand2'
        )
        
        # 使用place将按钮放在右下角
        copy_button.place(relx=1.0, rely=1.0, x=-15, y=-10, anchor='se')

        # 绑定点击和悬停事件
        def on_click(e):
            self.copy_template(template_text)
            copy_button.configure(bg='#D9D9D9')
            copy_button.after(100, lambda: copy_button.configure(bg='#EAEAEA'))
        
        def on_enter(e):
            copy_button.configure(bg='#D9D9D9')
        
        def on_leave(e):
            copy_button.configure(bg='#EAEAEA')
        
        copy_button.bind('<Button-1>', on_click)
        copy_button.bind('<Enter>', on_enter)
        copy_button.bind('<Leave>', on_leave)

    def copy_template(self, text_widget):
        """复制文本框内容到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text_widget.get('1.0', tk.END).strip())
        messagebox.showinfo("提示", "已复制到剪贴板！")

    def create_generate_button(self):
        button_frame = tk.Frame(self.main_frame, bg='#FFF0F5')
        button_frame.pack(pady=30)
        
        self.generate_button = ModernButton(
            button_frame,
            text="开始生成",
            command=self.generate_ppt,
            width=20,
            height=2,
            start_color='#FF2442',  # 小红书主色调
            end_color='#FF4D6D'
        )
        self.generate_button.pack()

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
            # 只设置文件夹路径
            self.save_path.set(dirname)

    def convert_ppt_to_images(self, ppt_path, batch_size=50):
        try:
            # 获取用户设置的尺寸
            width = int(self.width_var.get())
            height = int(self.height_var.get())
            
            # 获取文件名（不含扩展名）作为文件夹名
            base_name = os.path.splitext(os.path.basename(ppt_path))[0]
            
            # 创建与PPT同名的文件夹
            ppt_dir = os.path.dirname(ppt_path)
            images_dir = os.path.join(ppt_dir, base_name)
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)

            # 启动 WPS 应用程序
            wps = comtypes.client.CreateObject("KWPP.Application")
            wps.Visible = True
            
            try:
                # 打开PPT文件
                ppt = wps.Presentations
                presentation = ppt.Open(os.path.abspath(ppt_path))
                
                slide_count = presentation.Slides.Count
                for batch_start in range(0, slide_count, batch_size):
                    for i in range(batch_start, min(batch_start + batch_size, slide_count)):
                        slide = presentation.Slides.Item(i + 1)
                        # 生成输出文件路径
                        output_path = os.path.join(images_dir, f"{base_name}_第{i+1}页.jpg")
                        
                        # 导出当前幻灯片为JPG格式
                        slide.Export(output_path, "JPG", width, height)  # 添加宽度和高度参数
                        print(f"已将幻灯片 {i + 1} 保存为 {output_path}")
                        
                        # 使用Pillow确保图片质量和尺寸
                        img = Image.open(output_path)
                        img = img.resize((width, height), Image.Resampling.LANCZOS)  # 使用高质量的重采样方法
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
                
        except Exception as e:
            raise Exception(f"转换图片时出错: {str(e)}")

    def update_progress(self, value, message):
        """更新进度条和进度信息"""
        self.progress_var.set(value)
        self.progress_label.config(text=message)
        self.root.update()

    def generate_ppt(self):
        try:
            # 初始化进度
            self.update_progress(0, "开始处理...")
            
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
            save_dir = self.save_path.get()  # 现在这是文件夹路径
            if not os.path.exists(save_dir):
                messagebox.showerror("错误", "保存文件夹不存在")
                return

            # 生成完整的保存路径（添加文件名）
            current_time = time.strftime("%Y%m%d_%H%M%S")
            save_filename = f"小红书图文_{current_time}.pptx"
            full_save_path = os.path.join(save_dir, save_filename)

            self.update_progress(10, "读取Excel文件...")
            # 读取Excel文件，跳过空行，将 NA 值替换为空格
            df = pd.read_excel(self.excel_path.get()).dropna(how='all')
            # 将所有的 NA 值替换为空格
            df = df.fillna(' ')
            
            # 打印数据行数信息，用于调试
            print(f"总行数: {len(df)}")
            print(f"所有数据: {df.values.tolist()}")
            
            # 获取数据总行数
            total_rows = len(df)
            
            self.update_progress(20, "加载PPT模板...")
            try:
                # 创建 WPS 实例
                wps = comtypes.client.CreateObject("KWPP.Application")
                wps.Visible = True  # 需要保持True，否则可能出错
                
                # 最小化 WPS 窗口
                self.root.after(1000)
                
                # 查找 WPS 窗口并最小化
                def callback(hwnd, extra):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if 'WPS' in title or 'Presentation' in title:
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                
                win32gui.EnumWindows(callback, None)
                
                # 打开PPT文件
                ppt = wps.Presentations
                template = ppt.Open(os.path.abspath(self.ppt_path.get()))
                
                # 复制整个模板文件到新位置
                template.SaveAs(full_save_path)
                template.Close()  # 关闭模板文件
                
                # 打开新保存的文件进行编辑
                new_ppt = ppt.Open(full_save_path)
                
                # 获取单选按钮的值
                has_title = self.radio_var1.get() == "option1"
                unified_title = self.radio_var2.get() == "option2"
                
                # 获取第一页作为模板页（不删除它）
                template_slide = new_ppt.Slides(1)
                
                # 获取模板页面的背景属性
                template_background = template_slide.Background
                template_fill = template_background.Fill
                template_fore_color = template_fill.ForeColor.RGB
                template_back_color = template_fill.BackColor.RGB
                print(f"模板页面背景色信息:")
                print(f"- 填充类型: {template_fill.Type}")
                print(f"- 前景色: {template_fore_color}")
                print(f"- 背景色: {template_back_color}")
                old_font_size = None
                # 在处理每个形状之前，先保存模板页面的字号信息
                template_font_sizes = {}
                for shape in template_slide.Shapes:
                    try:
                        if shape.HasTextFrame:
                            template_font_sizes[shape.Name] = shape.TextFrame.TextRange.Font.Size
                    except:
                        continue

                # 遍历Excel的每一行数据（包括第一行）
                for i in range(len(df)):
                    progress = 20 + (i / total_rows * 40)
                    self.update_progress(progress, f"正在处理第 {i + 1} 行数据...")
                    
                    # 获取当前行数据
                    row = df.iloc[i]
                    
                    if i > 0:  # 第一页已经存在，只为后续数据创建新页面
                        # 复制第一页
                        new_ppt.Application.ActiveWindow.View.GotoSlide(1)  # 跳转到第一页
                        template_slide.Copy()  # 复制第一页
                        new_slide = new_ppt.Slides.Paste()  # 粘贴到末尾
                        
                        # 设置新页面的背景色，确保与模板一致
                        new_background = new_slide.Background
                        new_fill = new_background.Fill
                        new_fill.ForeColor.RGB = template_fore_color
                        new_fill.BackColor.RGB = template_back_color
                        
                        # 检查新页面的背景色
                        try:
                            print(f"\n第 {i+1} 页背景色信息:")
                            print(f"- 填充类型: {new_fill.Type}")
                            print(f"- 前景色: {new_fill.ForeColor.RGB}")
                            print(f"- 背景色: {new_fill.BackColor.RGB}")
                        except Exception as bg_error:
                            print(f"获取新页面背景色信息出错: {str(bg_error)}")
                    else:
                        # 使用第一页
                        new_slide = template_slide
                    
                    # 遍历所有形状并更新文本内容
                    for shape in new_slide.Shapes:
                        try:
                            if shape.HasTextFrame:
                                shape_name = shape.Name
                                print(f"处理形状: {shape_name}")
                                
                                if shape_name in df.columns:
                                    # 获取内容
                                    content = str(row[shape_name]).strip()
                                    if not content or content.lower() == 'nan':
                                        content = ' '
                                    
                                    # 从保存的模板中获取原始字号
                                    original_font_size = template_font_sizes.get(shape_name, shape.TextFrame.TextRange.Font.Size)
                                    print(f"原始字号: {original_font_size}")
                                    
                                    # 只有当内容包含"#我的首图#"时才设置字号
                                    if "#我的首图#" in content:
                                        print(f"检测到'#我的首图#'，设置字号为{self.font_size_var.get()}")
                                        shape.TextFrame.TextRange.Font.Size = int(self.font_size_var.get())
                                        content = content.replace("#我的首图#", "")
                                        print(f"已设置字号为{self.font_size_var.get()}，内容: {content}")
                                    else:
                                        # 其他内容使用模板中的原始字号
                                        shape.TextFrame.TextRange.Font.Size = original_font_size
                                        print(f"普通内容，使用原始字号: {original_font_size}")
                                    
                                    # 设置文本内容
                                    shape.TextFrame.TextRange.Text = content
                        except Exception as shape_error:
                            print(f"处理形状时出错: {str(shape_error)}")
                            continue
                
                self.update_progress(60, "保存PPT文件...")
                # 保存新的PPT文件
                new_ppt.SaveAs(full_save_path)
                
                self.update_progress(70, "转换为图片...")
                images_dir = self.convert_ppt_to_images(full_save_path)
                
                self.update_progress(90, "清理资源...")
                # 关闭文件和应用程序
                try:
                    new_ppt.Close()
                    wps.Quit()
                except:
                    pass
                
                self.update_progress(95, "打开生成的文件...")
                try:
                    os.startfile(full_save_path)
                    os.startfile(images_dir)
                except Exception as open_error:
                    print(f"打开文件失败: {str(open_error)}")
                
                self.update_progress(100, "处理完成！")
                messagebox.showinfo("成功", "PPT生成完成！图片已保存到images文件夹")
                
            finally:
                # 清理资源
                try:
                    if 'new_ppt' in locals():
                        new_ppt.Close()
                    if 'wps' in locals():
                        wps.Quit()
                except:
                    pass
                
        except Exception as e:
            self.update_progress(0, "处理出错")
            messagebox.showerror("错误", f"生成过程中出现错误：{str(e)}\n{traceback.format_exc()}")

def main():
    root = tk.Tk()
    app = PPTGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 