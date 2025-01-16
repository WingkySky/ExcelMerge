"""
合并配置模块
处理Excel合并相关的配置管理
"""
import tkinter as tk

class MergeConfig:
    def __init__(self):
        """初始化合并配置"""
        self.init_config()
        
    def init_config(self):
        """初始化所有配置变量"""
        # 合并模式设置
        self.merge_mode = tk.StringVar(value="single")  # single: 合并到单个sheet, multiple: 每个文件一个sheet
        self.sheet_name_mode = tk.StringVar(value="auto")  # auto: 使用文件名, original: 使用原sheet名, custom: 使用自定义名称
        self.custom_sheet_name = tk.StringVar(value="Sheet1")  # 自定义sheet名称
        
        # 数据区间设置
        self.start_row = tk.StringVar(value="1")
        self.end_row = tk.StringVar(value="")
        self.start_col = tk.StringVar(value="A")
        self.end_col = tk.StringVar(value="")
        
        # 表头设置
        self.header_row = tk.StringVar(value="1")  # 表头行号
        self.keep_header = tk.BooleanVar(value=True)  # 是否保留表头
        
        # 样式设置
        self.keep_styles = tk.BooleanVar(value=True)  # 是否保留样式
        self.keep_column_width = tk.BooleanVar(value=True)  # 是否保留列宽
        self.keep_cell_format = tk.BooleanVar(value=True)  # 是否保留单元格格式
        self.keep_colors = tk.BooleanVar(value=True)  # 是否保留颜色
        
    def get_merge_config(self):
        """获取合并配置字典"""
        return {
            'merge_mode': self.merge_mode.get(),
            'sheet_name_mode': self.sheet_name_mode.get(),
            'custom_sheet_name': self.custom_sheet_name.get(),
            'start_row': self.start_row.get(),
            'end_row': self.end_row.get(),
            'start_col': self.start_col.get(),
            'end_col': self.end_col.get(),
            'header_row': self.header_row.get(),
            'keep_header': self.keep_header.get(),
            'keep_styles': self.keep_styles.get(),
            'keep_column_width': self.keep_column_width.get(),
            'keep_cell_format': self.keep_cell_format.get(),
            'keep_colors': self.keep_colors.get()
        } 