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
        
        # 只保留一个样式选项
        option_frame = ctk.CTkFrame(style_frame)
        option_frame.pack(fill=tk.X, padx=5, pady=2)
        ctk.CTkCheckBox(option_frame, text="保留Excel原有样式", variable=self.app.merge_config.keep_styles,
                      **self.app.style_config.checkbox_style).pack(side=tk.LEFT, padx=20)
            
        # 添加说明文本
        note_frame = ctk.CTkFrame(self)
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        
        notes = [
            "样式设置说明：",
            "1. 勾选'保留Excel原有样式'将完整保留原Excel文件的所有样式设置",
            "2. 包括：字体、颜色、边框、对齐方式、数字格式等",
            "3. 不勾选则只保留原始数据，不保留任何样式"
        ]
        
        for note in notes:
            ctk.CTkLabel(note_frame, text=note, **self.app.style_config.label_style).pack(anchor=tk.W, padx=5, pady=2) 