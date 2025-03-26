import matplotlib
matplotlib.use('TkAgg')  # 添加这行在其他 import 之前
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import requests
from datetime import datetime, UTC
import re
import uuid
import tkinter
import ctypes  # 添加这行到文件顶部的导入部分
# 示例模块导入，替换为你的实际模块
import 指标整理
import log截图
import 文件名中字段批量修改
import 文件名操作
import 按照文件名匹配文件
from functools import lru_cache

# 本地文件路径
file_path = "deadline.txt"  # 使用txt后缀

# 下载文件的URL
url = "https://kdocs.cn/l/ceif9McrDxQ8"  # 替换为实际的文件URL

# 添加配置管理
CONFIG = {
    'UPDATE_INTERVAL_DAYS': 2,
    'WINDOW_SIZE': "700x650",
    'THEME_COLOR': "#f4f4f4",
    'BUTTON_COLOR': "#4CAF50",
}

# 添加日志记录功能
def setup_logging():
    """配置日志系统"""
    import logging
    logging.basicConfig(
        filename='app.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return logging.getLogger(__name__)

# 添加版本检查
def check_version():
    """检查软件版本"""
    current_version = "1.0.0"
    try:
        # 这里可以添加在线版本检查逻辑
        pass
    except Exception as e:
        print(f"版本检查失败：{e}")

def show_warning(message):
    """弹出告警对话框"""
    root = tkinter.Tk()
    root.withdraw()  # 隐藏主窗口
    messagebox.showerror("告警", message)  # 显示错误对话框
    root.destroy()

def download_txt_file():
    """下载txt文件并保存，并设置为隐藏状态"""
    try:
        response = requests.get(url, timeout=9)
        response.raise_for_status()  # 确保请求成功
        with open(file_path, "wb") as f:
            f.write(response.content)
        
        # 设置文件为隐藏属性
        FILE_ATTRIBUTE_HIDDEN = 0x02
        ret = ctypes.windll.kernel32.SetFileAttributesW(file_path, FILE_ATTRIBUTE_HIDDEN)
        if ret:
            print("文件下载成功并已设置为隐藏")
        else:
            print("文件下载成功，但设置隐藏属性失败")
    except requests.exceptions.RequestException as e:
        print(f"下载文件时出错：{e}")
        show_warning("网络有问题，请检查并重新运行软件。")
        exit()  # 退出程序
    except Exception as e:
        print(f"设置文件属性时出错：{e}")
    
def check_and_download_file():
    """检查文件是否存在，若不存在或需要更新则下载"""
    try:
        if not os.path.exists(file_path):
            print("文件不存在，正在下载...")
            download_txt_file()
            return
            
        # 文件存在，检查更新时间
        last_modified_time = os.path.getmtime(file_path)
        file_last_modified = datetime.fromtimestamp(last_modified_time, UTC)
        now = datetime.now(UTC)
        
        # 使用常量定义更新间隔
        UPDATE_INTERVAL_DAYS = 2
        if (now - file_last_modified).days > UPDATE_INTERVAL_DAYS:
            print("文件已过期，重新下载...")
            download_txt_file()
        else:
            print("文件已经是最新的，不需要重新下载。")
    except Exception as e:
        print(f"检查文件更新时出错：{e}")
        show_warning("检查更新失败，请检查网络连接。")

# 添加缓存装饰器
@lru_cache(maxsize=128)
def get_local_mac_address():
    """获取本机 MAC 地址（带缓存）"""
    mac = ":".join(re.findall('..', '%012x' % uuid.getnode()))
    print("本机 MAC 地址：", mac)
    return mac

def extract_and_check_authorization():
    """从 txt 文件中提取内容并检查授权状态"""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()

        # 使用更安全的正则表达式模式
        matches = re.findall(r'>2t2([^<]*)</text>', content)
        print("正在验证授权...")

        if not matches:
            print("未找到有效的授权内容。")
            show_warning("授权失败：未找到有效的授权内容。")
            return False

        mac_address = get_local_mac_address()
        for match in matches:
            match = match.strip()  # 清理空白字符
            if any(key in match for key in ['yesyes', mac_address]):
                print("授权验证成功")
                return True
            elif 'nono' in match:
                print("授权验证失败")
                show_warning("授权失败")
                return False

        print("授权验证失败")
        show_warning("授权失败")
        return False

    except Exception as e:
        print(f"授权验证出错：{e}")
        show_warning(f"授权验证时出错：{str(e)}")
        return False

# 主逻辑
check_and_download_file()  # 确保文件是最新的

try:
    extract_and_check_authorization()
except Exception as e:
    show_warning(f"程序运行异常：{e}")

# 创建主窗口
window = tk.Tk()
window.title("回单资料整理工具")
window.geometry("700x650")
window.config(bg="#f4f4f4")

# 全局变量来存储复选框状态
enable_4g5g_subgroup = tk.BooleanVar(value=True)

# 添加进度条
progress_label = tk.Label(window, text="完成进度提示", font=("Arial", 10), bg="#f4f4f4")
progress_label.pack(pady=5)

progress_bar = ttk.Progressbar(window, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(pady=10)

# 按钮样式
btn_style = {
    "font": ("Arial", 12),
    "width": 18,
    "height": 2,
    "bd": 3,
    "relief": "raised",
    "bg": "#4CAF50",
    "fg": "black",
    "activebackground": "#45a049",
    "activeforeground": "white",
}

# 按钮配置
column_buttons = [
    # 第一列
    [
        ("指标整理说明", lambda: show_instructions(1), {"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("指标整理", lambda: run_task(指标整理)),
        ("按照文件名匹配文件", lambda: run_in_thread(lambda: run_task(按照文件名匹配文件))),
    ],
    # 第二列
    [
        ("log截图说明", lambda: show_instructions(2), {"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("log截图", lambda: run_task(log截图)),
    ],
    # 第三列
    [
        ("文件名修改说明", lambda: show_instructions(3), {"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("文件名中字段批量修改", lambda: run_in_thread(lambda: run_task(文件名中字段批量修改))),
        ("文件名操作", lambda: run_in_thread(lambda: run_task(文件名操作))),
    ],
]

# 按钮布局
frame = tk.Frame(window, bg="#f4f4f4")
frame.pack(pady=10)

# 计算最大行数以对齐列
max_rows = max(len(col) for col in column_buttons)

# 创建网格布局
for col_idx, button_list in enumerate(column_buttons):
    for row_idx in range(max_rows):
        if row_idx < len(button_list):
            # 解包按钮配置，支持额外的样式
            text, command, *style_overrides = button_list[row_idx]
            # 合并默认样式和特定样式
            btn_style_updated = {**btn_style, **(style_overrides[0] if style_overrides else {})}
            # 创建按钮
            btn = tk.Button(frame, text=text, command=command, **btn_style_updated)
            btn.grid(row=row_idx, column=col_idx, padx=10, pady=5)
        else:
            # 添加空白标签占位以对齐列
            placeholder = tk.Label(frame, text="", bg="#f4f4f4", width=18, height=2)
            placeholder.grid(row=row_idx, column=col_idx, padx=10, pady=5)

# 添加4G/5G二次分组的复选框
subgroup_frame = tk.Frame(window, bg="#f4f4f4")
subgroup_frame.pack(pady=10)  # 增加上下间距

# 创建一个标签框来包含复选框
checkbox_labelframe = tk.LabelFrame(
    subgroup_frame,
    text="Log截图设置",
    font=("Arial", 10, "bold"),
    bg="#f4f4f4",
    fg="#333333",
    padx=10,
    pady=5,
    relief="groove",
    borderwidth=2
)
checkbox_labelframe.pack(padx=20)

# 美化复选框
subgroup_checkbox = tk.Checkbutton(
    checkbox_labelframe, 
    text="启用4G/5G二次分组", 
    variable=enable_4g5g_subgroup,
    font=("Arial", 10),
    bg="#f4f4f4",
    activebackground="#f0f0f0",
    activeforeground="#333333",
    selectcolor="#e0e0e0",
    cursor="hand2",  # 鼠标悬停时显示手型
    pady=3
)
subgroup_checkbox.pack(padx=10)

# 添加提示文本
help_text = tk.Label(
    checkbox_labelframe,
    text="勾选后将按4G/5G分别生成截图",
    font=("Arial", 9),
    fg="#666666",
    bg="#f4f4f4"
)
help_text.pack(pady=(0, 5))

# 按钮功能
def run_task(module):
    """运行指定模块的主函数"""
    try:
        update_progress(0)
        if not hasattr(module, 'main'):
            messagebox.showerror("错误", f"模块 {module.__name__} 缺少 main 函数")
            return

        if module.__name__ == 'log截图':
            window.after(100, lambda: safe_run_log_module(module))
        elif module.__name__ == '指标整理':
            window.after(100, lambda: safe_run_module(module))
        else:
            threading.Thread(target=lambda: safe_run_module(module)).start()

    except Exception as e:
        messagebox.showerror("运行错误", f"运行模块时出错：{str(e)}")
    finally:
        update_progress(100)

def safe_run_module(module):
    """在主线程中安全运行模块"""
    try:
        module.main()
        update_progress(100)
    except Exception as e:
        messagebox.showerror("运行错误", f"运行模块时出错：{e}")

def safe_run_log_module(module):
    """在主线程中安全运行log截图模块，传递复选框状态"""
    try:
        # 将复选框状态传递给log截图模块
        module.main(enable_4g5g_subgroup.get())
        update_progress(100)
    except Exception as e:
        messagebox.showerror("运行错误", f"运行模块时出错：{e}")

def update_progress(value):
    """更新进度条"""
    progress_bar["value"] = value
    window.update_idletasks()

def run_in_thread(func):
    """在后台线程中运行函数"""
    threading.Thread(target=func).start()

def show_instructions(number):
    """显示说明信息"""
    instructions = {
        1: """指标整理功能说明:

1. 功能概述:
   批量处理多个CSV指标文件，根据文件名自动分组并整理。

2. 使用步骤:
   - 点击"指标整理"按钮
   - 选择要处理的CSV文件(可多选)
   - 系统将自动分析文件名，按4G/5G分组
   - 处理结果会保存在程序当前目录

3. 特别说明:
   - 文件应包含"文件名称"列
   - 支持自动识别不同编码的CSV文件
   - 结果文件命名格式: [分组名称]指标.csv""",
        
        2: """Log截图功能说明:

1. 功能概述:
   可根据文件名生成分组截图，便于报告和展示。

2. 使用步骤:
   - 点击"log截图"按钮
   - 选择包含Log文件的文件夹
   - 系统按文件名自动分组并生成截图
   - 截图文件保存在程序当前目录

3. 4G/5G二次分组功能:
   - 勾选"启用4G/5G二次分组"时，系统会进一步将文件分为4G/5G两类
   - 不勾选时，仅按基础名称分组
   - 截图命名格式: [分组名称]-log截图.png""",
        
        3: """文件名修改功能说明:

1. 文件名中字段批量修改:
   - 用于批量替换文件名中的特定文本
   - 操作步骤:
     a. 选择要修改的文件
     b. 输入要替换的文本和替换后的文本
     c. 系统将自动处理所有选中文件

2. 文件名操作:
   - 通过Excel表格方式管理文件重命名
   - 操作步骤:
     a. 选择要重命名的文件
     b. 在Excel中编辑"目标文件名"列
     c. 保存Excel并返回程序
     d. 确认执行重命名操作

3. 注意事项:
   - 操作前请备份重要文件
   - Excel编辑时请确保保存更改
   - 使用前确认文件未被其他程序占用"""
    }
    
    if number in instructions:
        messagebox.showinfo(f"功能说明 {number}", instructions[number])

# 优化文件操作
def safe_file_operation(func):
    """文件操作装饰器"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except PermissionError:
            show_warning("文件访问被拒绝，请检查权限")
        except OSError as e:
            show_warning(f"文件操作失败：{e}")
    return wrapper

# 启动主循环
window.mainloop()