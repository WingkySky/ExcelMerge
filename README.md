# Excel文件合并工具

一个功能强大的Excel文件合并工具，支持多种合并方式和样式保留。

## 功能特点

- 支持多个Excel文件的合并
- 支持选择指定的Sheet和数据范围
- 支持保留原Excel文件的样式
- 支持多种合并模式：
  - 合并到单个Sheet
  - 每个文件单独一个Sheet
- 支持自定义Sheet名称
- 支持定时自动合并
- 界面美观，操作简单

## 系统要求

- Python 3.8 或更高版本
- Windows/macOS/Linux

## 安装说明

1. 克隆或下载本项目到本地
2. 安装依赖包：
```bash
pip install -r requirements.txt
```

## 使用说明

1. 运行程序：
```bash
python main.py
```

2. 主要功能：

### 文件选择
- 点击"添加文件"选择要合并的Excel文件
- 可以选择每个文件的具体Sheet
- 支持预览文件内容

### 合并设置
- 选择合并模式（单Sheet/多Sheet）
- 设置数据范围（起始行列）
- 设置表头选项

### Excel样式
- 可选择是否保留原文件样式
- 支持保留列宽、单元格格式、颜色等

### 定时任务
- 支持设置定时自动合并
- 可以选择每天固定时间执行

## 注意事项

1. 合并前请确保Excel文件未被其他程序占用
2. 建议在合并前预览数据，确保合并效果
3. 如果选择保留样式，合并时间可能会相对较长

## 项目结构

```
ExcelMerge/
├── main.py                 # 主程序入口
├── requirements.txt        # 依赖管理
├── README.md              # 项目说明
└── src/                   # 源代码目录
    ├── gui/               # GUI相关
    │   ├── main_window.py # 主窗口类
    │   ├── styles.py      # GUI样式配置
    │   └── dialogs.py     # 对话框相关
    ├── excel/             # Excel操作相关
    │   ├── merger.py      # Excel合并核心逻辑
    │   └── style_manager.py # Excel样式管理
    ├── utils/             # 工具类
    │   └── path_manager.py # 路径管理
    └── scheduler/         # 定时任务相关
        └── task_scheduler.py # 定时任务管理
```

## 开发说明

- 使用CustomTkinter构建现代化GUI界面
- 使用pandas进行Excel数据处理
- 使用openpyxl处理Excel样式
- 模块化设计，便于维护和扩展

## 版本历史

- v1.0.0 (2024-01-16)
  - 初始版本发布
  - 实现基本的Excel合并功能
  - 支持样式保留和定时任务

## 贡献指南

欢迎提交Issue和Pull Request来帮助改进这个项目。

## 许可证

本项目采用MIT许可证。详见LICENSE文件。
