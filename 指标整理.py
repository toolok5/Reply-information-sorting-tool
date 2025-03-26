import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

def get_group_name_from_file(file_name):
    """从文件名中提取分组名称（第一个字到4G/5G部分）"""
    # 查找 4G 或 5G 的位置
    g4_pos = file_name.find('4G')
    g5_pos = file_name.find('5G')
    
    # 确定结束位置
    if g4_pos != -1:
        end_pos = g4_pos + 2  # 包含 '4G'
    elif g5_pos != -1:
        end_pos = g5_pos + 2  # 包含 '5G'
    else:
        # 如果没有 4G 或 5G，返回整个文件名
        return file_name
    
    # 返回从开始到 4G/5G 的部分
    return file_name[:end_pos]

def process_csv_files():
    """处理CSV文件的主要逻辑"""
    # 弹出文件选择对话框，可多选csv文件
    file_paths = filedialog.askopenfilenames(
        title="请选择要处理的CSV文件",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    
    if not file_paths:  # 如果用户没有选择任何文件
        print("未选择文件，程序退出。")
        return
    
    # 用于存储所有分组的数据
    groups = {}
    
    # 修改获取当前目录的方式
    if getattr(sys, 'frozen', False):
        # 如果是打包后的程序
        current_dir = os.path.dirname(sys.executable)
    else:
        # 如果是开发环境
        current_dir = os.getcwd()
    
    # 逐个处理CSV文件
    for file_path in file_paths:
        try:
            # 首先尝试使用gbk编码读取
            try:
                df = pd.read_csv(file_path, encoding='gbk')
            except UnicodeDecodeError:
                # 如果gbk失败，尝试utf-8编码
                print(f"使用gbk编码读取失败，尝试使用utf-8编码: {os.path.basename(file_path)}")
                df = pd.read_csv(file_path, encoding='utf-8')
            
            if '文件名称' not in df.columns:
                print(f"警告: {os.path.basename(file_path)} 中没有'文件名称'列")
                continue
            
            # 处理每一行数据
            for _, row in df.iterrows():
                file_name = row['文件名称']
                if len(file_name) >= 5:
                    # 获取前5个字符作为初始分组
                    prefix = file_name[:5]
                    
                    # 确定4G/5G分组
                    if '4G' in file_name:
                        group_name = get_group_name_from_file(file_name)
                    elif '5G' in file_name:
                        group_name = get_group_name_from_file(file_name)
                    else:
                        group_name = prefix
                    
                    # 将数据添加到对应分组
                    if group_name not in groups:
                        groups[group_name] = []
                    groups[group_name].append(row)
            
            print(f"已处理文件: {os.path.basename(file_path)}")
            
        except Exception as e:
            error_msg = f"处理文件 {os.path.basename(file_path)} 时出错: {e}"
            print(error_msg)
            messagebox.showerror("错误", error_msg)
            continue
    
    # 保存每个分组到单独的CSV文件
    success_count = 0
    for group_name, rows in groups.items():
        if rows:  # 只处理非空分组
            output_file = os.path.join(current_dir, f"{group_name}指标.csv")
            try:
                # 创建DataFrame并保存为CSV，使用gbk编码
                df = pd.DataFrame(rows)
                df.to_csv(output_file, index=False, encoding='gbk')
                print(f"已保存分组 {group_name} 的数据到 {output_file}")
                success_count += 1
            except Exception as e:
                error_msg = f"保存文件 {output_file} 时出错: {e}"
                print(error_msg)
                messagebox.showerror("错误", error_msg)
    
    if success_count > 0:
        messagebox.showinfo("完成", f"处理完成！已成功保存 {success_count} 个分组的数据到 {current_dir} 目录。")

def main():
    """主函数"""
    try:
        process_csv_files()
    except Exception as e:
        messagebox.showerror("错误", f"程序运行出错：{e}")

if __name__ == "__main__":
    main()