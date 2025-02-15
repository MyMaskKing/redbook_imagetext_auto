# 小红书图文批量制作工具

一个用于批量生成小红书风格图文的工具，可以将 Excel 数据快速转换为 PPT 和图片。

## 功能特点

- 可视化界面操作，简单易用
- 支持批量将 Excel 数据导入 PPT
- 自动将 PPT 转换为高清图片
- 支持自定义图片尺寸
- 支持标题设置和处理选项

## 使用前提

- Windows 操作系统
- 必须安装 WPS 软件（用于 PPT 转图片功能）

## 使用说明

### 1. 准备文件

#### Excel 文件要求
- Excel 文件中的列名必须与 PPT 模板中的形状名称完全匹配
- 每一行数据将生成一页 PPT
- 第一行数据也会被处理（不作为表头）

#### PPT 模板要求
- PPT 中的文本框形状名称要与 Excel 的列名对应
- 建议使用 WPS 编辑 PPT 模板

### 2. 操作步骤

1. **选择文件**
   - 点击"选择"按钮选择 PPT 模板文件
   - 点击"选择"按钮选择 Excel 数据文件
   - 点击"选择"按钮选择保存位置

2. **设置选项**
   - 标题设置：选择"包含标题"或"只有正文"
   - 标题处理：选择"每页不同"或"统一标题"
   - 图片尺寸：设置导出图片的宽度和高度（默认 1920×1080）

3. **生成文件**
   - 点击"开始生成"按钮
   - 等待进度条完成
   - 程序会自动打开生成的文件和图片文件夹

### 3. 输出结果

程序会生成两种文件：
1. 一个新的 PPT 文件（包含所有数据）
2. 一个与 PPT 同名的文件夹，其中包含每页 PPT 转换的图片

## 注意事项

1. 确保 Excel 文件中的列名与 PPT 模板中的形状名称完全一致
2. 运行程序前必须安装 WPS 软件
3. 生成过程中请勿关闭弹出的 WPS 窗口
4. 建议使用较新版本的 WPS 软件
5. 如遇到错误，请检查文件路径是否包含特殊字符

## 常见问题

Q: 为什么必须安装 WPS？  
A: 程序使用 WPS 的接口来处理 PPT 和生成图片，这样可以保证最好的兼容性和输出质量。

Q: 生成的图片质量不够高？  
A: 可以在界面上调整图片尺寸，建议使用 1920×1080 或更高分辨率。

Q: Excel 数据没有正确写入 PPT？  
A: 请确保 Excel 的列名与 PPT 中的形状名称完全匹配，包括空格和大小写。

## 技术支持

如果遇到问题，请检查：
1. WPS 是否正确安装
2. Excel 文件格式是否正确
3. PPT 模板中的形状名称是否正确
4. 文件路径是否包含特殊字符

## AI提问模板
请帮我查找关于“2025年手机发布”的内容，生成的格式为表格，有三列：标题，内容，预测价格,并且帮我生成150字的小红书爆文，要求足够吸引人眼球，里面可以插入一些表情。

## 版本历史

v1.0.0
- 初始发布
- 支持基本的 Excel 到 PPT 转换功能
- 支持 PPT 到图片的转换
- 提供可视化界面 