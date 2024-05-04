import os
import opencc
import tkinter as tk
from tkinter import filedialog
from docx import Document
import threading
from tkinter import messagebox

# 定义要遍历的文件夹路径
root_folder = ''  # 初始化为空字符串

# 定义一个函数以打开文件对话框并获取文件夹路径
def get_folder():
    global root_folder
    folder_path = filedialog.askdirectory()
    if folder_path:
        root_folder = folder_path
        print("所选文件夹：", root_folder)
        if (root_folder):
            # 显示所选文件夹
            label_selected_folder.config(text=root_folder)
            label.config(text="")
            # 简转台繁按钮状态设置为可用
            button_convert_s2tw.config(state="normal")
            # 台繁转简按钮状态设置为可用
            button_convert_tw2s.config(state="normal")
            # 简转港繁按钮状态设置为可用
            button_convert_s2hk.config(state="normal")
            # 港繁转简按钮状态设置为可用
            button_convert_hk2s.config(state="normal")

# 新线程内执行转换
def convert_folder_thread(conversion):
    t = threading.Thread(target=convert_folder, args=(conversion,))
    t.start()

# 定义一个函数以遍历文件夹及子文件夹，并对docx文件进行简体转繁体转换
def convert_folder(conversion):
    # 弹窗提醒和确认
    user_confirmation = messagebox.askyesno("确认操作", 
                                            "该操作会遍历所选文件夹及其子文件夹，并对其中的docx文件进行简繁转换，转换后的文件会覆盖原文件，请确认是否继续？",
                                            parent=root)
    if not user_confirmation:
        return;

    # 简转台繁和台繁转简按钮状态设置为不可用
    button_convert_s2tw.config(state="disabled")  
    button_convert_tw2s.config(state="disabled")
    # 简转港繁和港繁转简按钮状态设置为不可用
    button_convert_s2hk.config(state="disabled")
    button_convert_hk2s.config(state="disabled")

    # 创建OpenCC对象以简繁转换
    converter = opencc.OpenCC(conversion)

    # 使用os.walk()遍历文件夹及子文件夹
    for folder_name, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            # 检查文件名是否以.docx结尾
            if filename.endswith('.docx'):
                # 构建docx文件的完整路径
                docx_file_path = os.path.join(folder_name, filename)
                # 读取简体docx文件
                doc = Document(docx_file_path)

                # 遍历文档段落并替换为繁体中文
                for para in doc.paragraphs:
                    if para.text.strip():  # 只处理非空文本
                        para.text = converter.convert(para.text)

                # 保存为繁体docx文件（也可以覆盖原始文件）
                doc.save(docx_file_path)

                # 获取新的文件名并将文件改名为繁体字名称
                new_filename = converter.convert(filename)
                new_docx_file_path = os.path.join(folder_name, new_filename)
                os.rename(docx_file_path, new_docx_file_path)

                # 更新显示
                sub_filename = filename[:7]
                label_selected_folder.config(text=sub_filename + "..." )
                
                print("转换并保存：", docx_file_path)
    
    # 完成转换
    label_selected_folder.config(text="转换完成！")

# 创建主窗口
root = tk.Tk()
root.title("S2T简繁转换")

# 创建一个容器
container = tk.Frame(root)
container.pack(pady=260)

# 创建一个标签
label = tk.Label(container, text="请选择一个文件夹：")
label.pack(side="left")  # 设置标签在左侧显示

# 创建一个标签以显示所选文件夹
label_selected_folder = tk.Label(container, text="")
label_selected_folder.pack(side="left")  # 设置标签在左侧显示

# 创建一个按钮以打开文件对话框
button = tk.Button(container, text="浏览", command=get_folder)
button.pack(side="left")  # 设置按钮在左侧显示，并添加水平间距

# 简转台繁按钮
button_convert_s2tw = tk.Button(container, text="简转台繁", command=lambda: convert_folder_thread("s2tw"))
button_convert_s2tw.pack(side="left", padx=10)  # 设置按钮在左侧显示，并添加水平间距
button_convert_s2tw.config(state="disabled")  # 按钮默认状态为禁用

# 台繁转简按钮
button_convert_tw2s = tk.Button(container, text="台繁转简", command=lambda: convert_folder_thread("tw2s"))
button_convert_tw2s.pack(side="left")  # 设置按钮在左侧显示
button_convert_tw2s.config(state="disabled")  # 按钮默认状态为禁用

# 简转港繁按钮
button_convert_s2hk = tk.Button(container, text="简转港繁", command=lambda: convert_folder_thread("s2hk"))
button_convert_s2hk.pack(side="left")  # 设置按钮在左侧显示
button_convert_s2hk.config(state="disabled")  # 按钮默认状态为禁用

# 港繁转简按钮
button_convert_hk2s = tk.Button(container, text="港繁转简", command=lambda: convert_folder_thread("hk2s"))
button_convert_hk2s.pack(side="left")  # 设置按钮在左侧显示
button_convert_hk2s.config(state="disabled")  # 按钮默认状态为禁用

# 主窗口屏幕正中显示
# 计算屏幕尺寸
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 计算窗口尺寸
window_width = 800
window_height = 600

# 计算窗口位置
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# 设置窗口位置
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 禁止调整窗口大小
root.resizable(width=False, height=False) 

# 运行 GUI 事件循环
root.mainloop()