# 回单资料整理工具

## 功能说明

### 1. 指标整理
- 支持批量处理指标文件
- 自动整理并生成结果文件
- 支持多文件同时处理
- 使用matplotlib进行数据可视化展示

### 2. Log截图
- 支持按文件名分组生成截图
- 可选择是否启用4G/5G二次分组功能
- 自动生成整理后的截图文件

### 3. 文件名修改功能
#### 3.1 文件名中字段批量修改
- 支持批量修改文件名中的指定字段
- 通过简单的输入实现批量替换

#### 3.2 文件名操作
- 支持通过Excel表格方式管理文件重命名
- 可同时处理多个文件
- 提供直观的Excel界面进行编辑
- 支持预览和确认更改

## 技术实现
- Python 3.x
- 核心依赖库：
  - matplotlib: 数据可视化
  - pandas: 数据处理
  - openpyxl: Excel文件操作
  - pillow: 图像处理

## 使用说明

### Log截图功能
1. 选择包含log文件的文件夹
2. 可选择是否启用4G/5G二次分组
   - 启用：将按4G/5G分别生成截图
   - 不启用：仅按基础分组生成截图
3. 自动在当前目录生成截图文件

### 文件名修改功能
#### 字段批量修改
1. 选择需要修改的文件
2. 输入要替换的字段（格式：原字段,新字段）
3. 确认后自动完成修改

#### 文件名操作
1. 选择需要重命名的文件
2. 在自动打开的Excel文件中编辑目标文件名
3. 保存Excel文件后返回程序
4. 确认是否执行重命名操作

## 系统要求
- Windows 10及以上
- Python 3.7+
- 2GB以上可用内存
- 500MB以上可用磁盘空间

## 安装说明
1. 克隆仓库
```bash
git clone https://github.com/toolok5/Reply-information-sorting-tool.git
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 运行程序
```bash
python main.py
```

## 注意事项
- 使用前请确保相关文件未被其他程序占用
- Excel操作时请及时保存更改
- 建议在操作前备份重要文件
- 首次运行时会自动安装所需依赖库

## 联系方式
如有问题欢迎联系，微信手机号：15057337780

## 授权说明
本软件需要授权才能使用，首次运行时会自动进行授权验证。

## 更新日志
### v1.0.0 (当前版本)
- 初始版本发布
- 实现基础的指标整理功能
- 实现Log截图功能
- 实现文件名修改功能