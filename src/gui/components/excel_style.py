"""
Excel样式设置组件模块
处理Excel样式设置界面的创建和交互
"""
import tkinter as tk
import customtkinter as ctk

class ExcelStyle(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化Excel样式设置器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="Excel文件样式设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        # Excel样式选项
        style_frame = ctk.CTkFrame(self)
        style_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 创建样式选项
        style_options = [
            ("保留Excel原有样式", self.app.merge_config.keep_styles),
            ("保留列宽", self.app.merge_config.keep_column_width),
            ("保留单元格格式", self.app.merge_config.keep_cell_format),
            ("保留颜色", self.app.merge_config.keep_colors)
        ]
        
        for text, var in style_options:
            option_frame = ctk.CTkFrame(style_frame)
            option_frame.pack(fill=tk.X, padx=5, pady=2)
            ctk.CTkCheckBox(option_frame, text=text, variable=var,
                          **self.app.style_config.checkbox_style).pack(side=tk.LEFT, padx=20)
            
        # 添加说明文本
        note_frame = ctk.CTkFrame(self)
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        
        notes = [
            "样式设置说明：",
            "1. 保留Excel原有样式：包括字体、边框、对齐方式等",
            "2. 保留列宽：保持原Excel文件的列宽设置",
            "3. 保留单元格格式：保持数字、日期等格式设置",
            "4. 保留颜色：包括背景色和字体颜色"
        ]
        
        for note in notes:
            ctk.CTkLabel(note_frame, text=note, **self.app.style_config.label_style).pack(anchor=tk.W, padx=5, pady=2) 