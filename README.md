# Excel文件合并工具

一个功能强大的Excel文件合并工具，支持多种合并模式和样式保留选项。

## 功能特点

### 1. 文件选择
- 支持选择多个Excel文件（.xlsx, .xls）
- 可以选择每个文件的指定Sheet
- 提供文件列表预览和管理
- 支持快速路径选择和自定义输出路径

### 2. 合并模式
- 单Sheet模式：将所有数据合并到一个Sheet
- 多Sheet模式：每个文件保存为单独的Sheet
- 支持自定义Sheet命名规则：
  - 使用文件名
  - 使用原Sheet名
  - 自定义名称

### 3. 数据选择
- 可指定数据范围（起始行/列、结束行/列）
- 支持表头设置
- 可选择是否保留表头

### 4. 样式设置
- 保留Excel原有样式（字体、边框、对齐等）
- 保留列宽设置
- 保留单元格格式（数字、日期等）
- 保留颜色设置（背景色、字体颜色）

### 5. 界面设置
- 支持多种外观模式：
  - 跟随系统
  - 浅色模式
  - 深色模式
- 支持多种颜色主题：
  - Blue
  - Dark Blue
  - Green

### 6. 定时任务
- 支持设置定时执行合并任务
- 24小时制时间设置
- 每日自动执行
- 可随时启动/停止

## 安装说明

1. 确保已安装Python 3.8或更高版本
2. 克隆项目到本地：
```bash
git clone https://github.com/yourusername/ExcelMerge.git
cd ExcelMerge
```

3. 创建并激活虚拟环境：
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

4. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用说明

1. 启动程序：
```bash
python main.py
```

2. 基本使用流程：
   - 在"文件选择"页面添加需要合并的Excel文件
   - 选择合适的合并模式和设置
   - 配置样式保留选项
   - 选择输出路径
   - 点击"立即执行合并"或设置定时任务

3. 数据预览：
   - 可以预览选中的单个文件
   - 可以预览合并后的结果

## 注意事项

1. 合并前请确保：
   - 所有文件格式正确
   - 数据结构相似（尤其是合并到单个Sheet时）
   - 有足够的磁盘空间

2. 定时任务注意：
   - 设置定时任务后请勿关闭程序
   - 确保计算机不会进入休眠状态

## 系统要求

- 操作系统：Windows/macOS/Linux
- Python版本：3.8+
- 内存：4GB及以上（建议）
- 磁盘空间：根据处理的Excel文件大小决定

## 依赖库

- customtkinter
- pandas
- openpyxl
- tkinter (Python标准库)

## 许可证

MIT License
