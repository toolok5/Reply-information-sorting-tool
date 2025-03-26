import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import time
import shutil

# 全局变量，用于存储选择的文件及其路径
selected_files = []
file_paths = {}

def select_files():
    """打开文件选择对话框，让用户选择多个文件"""
    global selected_files, file_paths
    
    filetypes = [
        ('CSV files', '*.csv'),
        ('Excel files', '*.xlsx;*.xls'),
        ('Image files', '*.png;*.jpg;*.jpeg'),
        ('All files', '*.*')
    ]
    files = filedialog.askopenfilenames(title="选择文件", filetypes=filetypes)
    
    if files:
        selected_files = list(files)
        # 存储文件路径信息
        for file_path in files:
            filename = os.path.basename(file_path)
            file_paths[filename] = file_path
    
    return files if files else []

def write_to_excel(files, excel_path):
    """将选择的文件名写入Excel文件的'原文件名'列"""
    # 检查Excel文件是否存在
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            # 确保必要的列存在
            if "原文件名" not in df.columns:
                df["原文件名"] = ""
            if "目标文件名" not in df.columns:
                df["目标文件名"] = ""
        except Exception as e:
            messagebox.showerror("错误", f"无法读取Excel文件: {e}\n请确保Excel文件未被其他程序打开")
            return False
    else:
        # 创建新的Excel文件
        df = pd.DataFrame(columns=["原文件名", "目标文件名"])
    
    # 将文件名添加到DataFrame，避免重复
    new_entries = []
    for file in files:
        filename = os.path.basename(file)
        if filename not in df["原文件名"].values:
            new_entries.append({"原文件名": filename, "目标文件名": ""})
    
    if new_entries:
        df = pd.concat([df, pd.DataFrame(new_entries)], ignore_index=True)
        # 保存DataFrame到Excel
        try:
            df.to_excel(excel_path, index=False)
            return True
        except Exception as e:
            messagebox.showerror("错误", f"无法保存Excel文件: {e}\n请确保Excel文件未被其他程序打开")
            return False
    else:
        messagebox.showinfo("提示", "所选文件已全部存在于Excel中")
        return True

def show_prompt():
    """显示提示框，包含'下一步'和'结束'按钮"""
    result = messagebox.askyesno("操作", "请确保已保存Excel中的更改\n\n点击'是'将原文件重命名为目标文件名\n点击'否'结束程序")
    return result

def process_next_step(excel_path):
    """处理'下一步'按钮的逻辑，将原文件重命名为目标文件名列中的名称"""
    global file_paths
    
    try:
        # 尝试多次读取Excel文件，以防文件被锁定
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                df = pd.read_excel(excel_path)
                break
            except Exception as e:
                if attempt < max_attempts - 1:
                    messagebox.showwarning("警告", f"无法读取Excel文件，正在重试... ({attempt+1}/{max_attempts})")
                    time.sleep(1)
                else:
                    messagebox.showerror("错误", f"无法读取Excel文件: {e}\n请确保Excel文件未被其他程序打开")
                    return False
        
        # 检查是否有空的目标文件名
        if df["目标文件名"].isna().any() or (df["目标文件名"] == "").any():
            if not messagebox.askyesno("警告", "有些目标文件名为空，是否继续？"):
                return False
        
        # 重命名文件
        success_count = 0
        error_count = 0
        error_messages = []
        
        for _, row in df.iterrows():
            source_name = row["原文件名"]
            target_name = str(row["目标文件名"])  # 确保目标文件名是字符串
            
            # 跳过空的目标文件名
            if pd.isna(target_name) or target_name == "":
                continue
            
            # 检查是否有源文件路径信息
            if source_name not in file_paths:
                error_count += 1
                error_messages.append(f"找不到源文件路径信息: {source_name}")
                continue
                
            source_path = file_paths[source_name]
            source_ext = os.path.splitext(source_path)[1]  # 获取源文件扩展名
            
            # 如果目标文件名没有扩展名，则添加源文件的扩展名
            if not os.path.splitext(target_name)[1]:
                target_name = target_name + source_ext
            
            target_dir = os.path.dirname(source_path)
            target_path = os.path.join(target_dir, target_name)
            
            # 检查源文件是否存在
            if not os.path.exists(source_path):
                error_count += 1
                error_messages.append(f"找不到源文件: {source_path}")
                continue
                
            # 检查目标文件是否已存在
            if os.path.exists(target_path) and source_path != target_path:
                if not messagebox.askyesno("确认", f"文件 {target_name} 已存在，是否覆盖？"):
                    continue
                    
            try:
                # 重命名文件
                shutil.move(source_path, target_path)
                # 更新文件路径信息
                file_paths[target_name] = target_path
                del file_paths[source_name]
                success_count += 1
            except Exception as e:
                error_count += 1
                error_messages.append(f"重命名 {source_name} 到 {target_name} 失败: {e}")
        
        # 显示结果
        result_message = f"操作完成\n成功: {success_count} 个文件"
        if error_count > 0:
            result_message += f"\n失败: {error_count} 个文件"
            for msg in error_messages[:5]:  # 只显示前5个错误
                result_message += f"\n- {msg}"
            if len(error_messages) > 5:
                result_message += f"\n... 还有 {len(error_messages) - 5} 个错误未显示"
                
        messagebox.showinfo("结果", result_message)
        return True
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时出错: {e}")
        return False

def open_excel_file(excel_path):
    """尝试打开Excel文件供用户编辑"""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(excel_path)
        elif os.name == 'posix':  # macOS or Linux
            if os.path.exists('/usr/bin/open'):  # macOS
                subprocess.run(['open', excel_path])
            else:  # Linux
                subprocess.run(['xdg-open', excel_path])
        return True
    except Exception as e:
        messagebox.showwarning("警告", f"无法自动打开Excel文件: {e}\n请手动打开并编辑文件: {excel_path}")
        return False

def main():
    """主函数"""
    # 创建根窗口但不显示
    root = tk.Tk()
    root.withdraw()
    
    # 定义Excel文件路径
    excel_filename = "文件名操作.xlsx"
    excel_path = os.path.join(os.getcwd(), excel_filename)
    
    # 选择文件
    files = select_files()
    if not files:
        messagebox.showinfo("提示", "未选择任何文件，程序结束")
        return
    
    # 写入Excel
    success = write_to_excel(files, excel_path)
    if not success:
        return
    
    messagebox.showinfo("成功", f"已将{len(files)}个文件名写入Excel文件\n请在Excel中编辑'目标文件名'列，完成后保存Excel文件")
    
    # 尝试打开Excel文件供用户编辑
    open_excel_file(excel_path)
    
    # 提示用户下一步操作
    while True:
        # 显示提示框
        next_step = show_prompt()
        if next_step:  # 用户点击"是"（下一步）
            success = process_next_step(excel_path)
            if not success:
                break
        else:  # 用户点击"否"（结束）
            messagebox.showinfo("提示", "程序结束")
            break

if __name__ == "__main__":
    main()