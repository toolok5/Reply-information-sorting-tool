import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

def select_excel_file():
    """选择Excel文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择包含文件名的Excel表格",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    root.destroy()  # 确保窗口被销毁
    return file_path

def select_source_folder():
    """选择源文件夹"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="选择要扫描的源文件夹")
    root.destroy()  # 确保窗口被销毁
    return folder_path

def create_result_folder(source_folder):
    """创建结果文件夹"""
    result_folder = os.path.join(source_folder, "匹配结果")
    os.makedirs(result_folder, exist_ok=True)
    return result_folder

def read_excel_filenames(excel_path, column_index=0):
    """读取Excel中的文件名列表"""
    try:
        df = pd.read_excel(excel_path)
        # 获取指定列的所有值，默认使用第一列
        filenames = df.iloc[:, column_index].astype(str).tolist()
        # 过滤掉空值和NaN
        filenames = [name for name in filenames if name and name.lower() != 'nan']
        return filenames
    except Exception as e:
        messagebox.showerror("错误", f"读取Excel文件时出错: {str(e)}")
        return []

def match_and_move_all_items(source_folder, result_folder, filenames):
    """匹配并移动所有类型的文件和文件夹"""
    moved_count = 0
    matched_items = []
    
    # 获取源文件夹中的所有项目（文件和文件夹）
    for item in os.listdir(source_folder):
        item_path = os.path.join(source_folder, item)
        
        # 跳过结果文件夹
        if os.path.abspath(item_path) == os.path.abspath(result_folder):
            continue
            
        # 检查项目名是否包含任何目标文件名
        for filename in filenames:
            if filename in item:
                matched_items.append((item, filename))
                # 构建目标路径
                dest_path = os.path.join(result_folder, item)
                
                try:
                    # 移动文件或文件夹
                    shutil.move(item_path, dest_path)
                    moved_count += 1
                    break  # 一旦找到匹配项就跳出内循环
                except Exception as e:
                    messagebox.showwarning("警告", f"移动项目 '{item}' 时出错: {str(e)}")
    
    return moved_count, matched_items

def select_excel_column():
    """选择Excel中的列，使用简单的对话框"""
    root = tk.Tk()
    root.title("选择列")
    root.geometry("300x150")
    root.attributes('-topmost', True)  # 确保窗口在最前面
    
    selected_column = tk.StringVar(value="1")
    
    tk.Label(root, text="请输入包含文件名的列号(从1开始):", pady=10).pack()
    entry = tk.Entry(root, textvariable=selected_column)
    entry.pack(pady=10)
    
    result = [0]  # 使用列表存储结果，以便在函数内部修改
    
    def on_submit():
        try:
            col_num = int(selected_column.get())
            if col_num < 1:
                raise ValueError("列号必须大于0")
            result[0] = col_num - 1  # 转换为0-based索引
            root.quit()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的列号")
    
    tk.Button(root, text="确定", command=on_submit).pack(pady=10)
    
    # 确保窗口关闭时也能退出主循环
    root.protocol("WM_DELETE_WINDOW", root.quit)
    
    root.mainloop()
    root.destroy()
    
    return result[0]

def main():
    print("程序开始运行...")
    
    # 选择Excel文件
    print("请选择Excel文件...")
    excel_path = select_excel_file()
    if not excel_path:
        messagebox.showinfo("提示", "未选择Excel文件，程序已取消")
        return
    print(f"已选择Excel文件: {excel_path}")
    
    # 选择Excel列
    print("请选择Excel列...")
    column_index = select_excel_column()
    print(f"已选择列索引: {column_index}")
    
    # 读取文件名列表
    print("正在读取文件名列表...")
    filenames = read_excel_filenames(excel_path, column_index)
    if not filenames:
        messagebox.showinfo("提示", "未找到有效的文件名，程序已取消")
        return
    print(f"读取到 {len(filenames)} 个文件名")
    
    # 选择源文件夹
    print("请选择源文件夹...")
    source_folder = select_source_folder()
    if not source_folder:
        messagebox.showinfo("提示", "未选择源文件夹，程序已取消")
        return
    print(f"已选择源文件夹: {source_folder}")
    
    # 创建结果文件夹
    print("正在创建结果文件夹...")
    result_folder = create_result_folder(source_folder)
    print(f"已创建结果文件夹: {result_folder}")
    
    # 匹配并移动所有项目
    print("正在匹配并移动所有文件和文件夹...")
    moved_count, matched_items = match_and_move_all_items(source_folder, result_folder, filenames)
    
    # 显示结果
    if moved_count > 0:
        result_message = f"成功移动了 {moved_count} 个匹配的项目到 '{os.path.basename(result_folder)}' 文件夹。\n\n匹配详情:\n"
        for item, pattern in matched_items:
            result_message += f"项目: {item} 匹配: {pattern}\n"
        messagebox.showinfo("完成", result_message)
    else:
        messagebox.showinfo("完成", "未找到匹配的项目")
    
    print("程序运行完成")

if __name__ == "__main__":
    main()