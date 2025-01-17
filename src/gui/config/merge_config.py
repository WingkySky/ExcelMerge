"""
合并配置模块
处理Excel合并相关的配置管理
"""
import tkinter as tk

class MergeConfig:
    def __init__(self):
        """初始化合并配置"""
        # 合并模式
        self.merge_mode = tk.StringVar(value="single")  # single: 合并到单个sheet, multiple: 每个文件一个sheet
        self.sheet_name_mode = tk.StringVar(value="auto")  # auto: 使用文件名, original: 使用原sheet名, custom: 使用自定义名称
        self.custom_sheet_name = tk.StringVar(value="Sheet1")  # 自定义sheet名称
        
        # 数据范围
        self.start_row = tk.StringVar(value="1")
        self.end_row = tk.StringVar(value="")
        self.start_col = tk.StringVar(value="A")
        self.end_col = tk.StringVar(value="")
        
        # 表头设置
        self.header_row = tk.StringVar(value="1")  # 表头行号
        self.keep_header = tk.BooleanVar(value=True)  # 是否保留表头
        
        # 样式设置（简化为单个选项）
        self.keep_styles = tk.BooleanVar(value=True)  # 是否保留所有样式
        
    def get_merge_config(self):
        """获取合并配置"""
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
            'keep_styles': self.keep_styles.get()
        } 