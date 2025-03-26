import os
from datetime import datetime
import matplotlib.pyplot as plt # type: ignore
from matplotlib.font_manager import FontProperties
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox

def get_group_name_from_file(file_name):
    """从文件名中提取分组名称（第一个字到4G/5G前的字符）"""
    # 查找 4G 或 5G 的位置
    g4_pos = file_name.find('4G')
    g5_pos = file_name.find('5G')
    
    # 确定结束位置
    if g4_pos != -1:
        end_pos = g4_pos  # 不包含 '4G'，到4G前面一个字符
    elif g5_pos != -1:
        end_pos = g5_pos  # 不包含 '5G'，到5G前面一个字符
    else:
        # 如果没有 4G 或 5G，返回整个文件名
        return file_name
    
    # 返回从开始到 4G/5G 前的部分
    return file_name[:end_pos]

def capture_files_group(folder_path, enable_4g5g_subgroup=True):
    # 获取文件夹中的所有文件
    files = sorted(os.listdir(folder_path))
    
    # 按照文件名中从第一个字符到4G/5G前的内容进行分组
    groups = {}
    for file in files:
        # 使用get_group_name_from_file函数获取基础分组名称
        base_group_name = get_group_name_from_file(file)
        
        if base_group_name not in groups:
            groups[base_group_name] = []
            
        # 获取文件信息
        file_path = os.path.join(folder_path, file)
        size_kb = os.path.getsize(file_path) / 1024
        mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
        mod_time_str = mod_time.strftime("%Y/%m/%d %H:%M")
        file_type = os.path.splitext(file)[1][1:].upper() + " 文件"
        
        groups[base_group_name].append({
            'name': file,
            'date': mod_time_str,
            'type': file_type,
            'size': f"{size_kb:,.0f} KB"
        })
    
    # 确定是否需要进行4G/5G二次分组
    final_groups = {}
    if enable_4g5g_subgroup:
        # 进行4G/5G二次分组
        for group_name, files_list in groups.items():
            # 创建4G和5G的子分组
            g4_files = []
            g5_files = []
            other_files = []
            
            for file_info in files_list:
                if '4G' in file_info['name']:
                    g4_files.append(file_info)
                elif '5G' in file_info['name']:
                    g5_files.append(file_info)
                else:
                    other_files.append(file_info)
            
            # 只添加非空的分组
            if g4_files:
                final_groups[f"{group_name}4G"] = g4_files
            if g5_files:
                final_groups[f"{group_name}5G"] = g5_files
            if other_files:
                final_groups[group_name] = other_files
    else:
        # 不进行二次分组，直接使用原始分组
        final_groups = groups
    
    # 为每个最终分组创建截图
    for group_name, group_files in final_groups.items():
        plt.figure(figsize=(15, max(3, (len(group_files) + 1) * 0.6)))
        plt.axis('off')
        
        plt.gca().set_facecolor('white')
        plt.gcf().set_facecolor('white')
        
        # 定义统一的起始位置
        start_x = 0.02
        
        # 添加表头
        headers = ["名称", "修改日期", "类型", "大小"]
        header_positions = [start_x, 0.45, 0.65, 0.85]
        
        # 添加表头
        for i, (header, pos) in enumerate(zip(headers, header_positions)):
            plt.text(pos, 0.95, header, fontsize=10, fontfamily='SimHei',
                    verticalalignment='center', transform=plt.gca().transAxes,
                    color='black')
            
            if i < len(headers) - 1:
                next_pos = header_positions[i + 1] - 0.02
                plt.plot([next_pos, next_pos], [0.90, 0.98], 
                        color='#E5E5E5', linewidth=1, transform=plt.gca().transAxes)
        
        # 添加文件列表
        start_y = 0.80
        
        # 设置支持 Emoji 的字体
        try:
            emoji_font = FontProperties(family='Segoe UI Emoji')
        except:
            emoji_fonts = ['Segoe UI Symbol', 'Apple Color Emoji', 'Noto Color Emoji', 'Noto Emoji']
            for font in emoji_fonts:
                try:
                    emoji_font = FontProperties(family=font)
                    break
                except:
                    continue
        
        for i, file_info in enumerate(group_files):
            y_pos = start_y - (i * 0.1)
            
            # 使用文本图标
            plt.text(start_x, y_pos, "📄 ", fontproperties=emoji_font,
                    fontsize=10, transform=plt.gca().transAxes)
            
            # 添加文件信息
            plt.text(start_x + 0.02, y_pos, file_info['name'], 
                    fontsize=10, fontfamily='SimHei',
                    transform=plt.gca().transAxes)
            plt.text(0.45, y_pos, file_info['date'], 
                    fontsize=10, fontfamily='SimHei',
                    transform=plt.gca().transAxes)
            plt.text(0.65, y_pos, file_info['type'], 
                    fontsize=10, fontfamily='SimHei',
                    transform=plt.gca().transAxes)
            plt.text(0.95, y_pos, file_info['size'], 
                    fontsize=10, fontfamily='SimHei',
                    transform=plt.gca().transAxes,
                    horizontalalignment='right')
        
        # 修改输出文件名的生成方式
        if group_files:
            if enable_4g5g_subgroup:
                output_file = f"{group_name}-log截图.png"
            else:
                # 不进行二次分组时，使用基础分组名称
                output_file = f"{group_name}-log截图.png"
            
            plt.savefig(output_file, bbox_inches='tight', dpi=500, 
                       facecolor='white', edgecolor='none')
            plt.close()
            
            print(f"已保存分组截图到 {output_file}")

def main(enable_4g5g_subgroup=True):
    """主函数，增加参数控制是否进行4G/5G二次分组"""
    try:
        # 获取文件夹路径
        folder_path = filedialog.askdirectory(title="请选择要处理的文件夹")
        
        if folder_path:  # 如果用户选择了文件夹（而不是取消）
            capture_files_group(folder_path, enable_4g5g_subgroup)
            messagebox.showinfo("完成", "文件处理完成！")
        else:
            print("未选择文件夹，程序退出。")
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出错：{e}")

if __name__ == "__main__":
    main()