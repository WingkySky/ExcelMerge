"""
底部框架组件模块
处理底部按钮和状态栏的创建和交互
"""
import tkinter as tk
import customtkinter as ctk

class BottomFrame(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化底部框架"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 预览按钮
        preview_frame = ctk.CTkFrame(self)
        preview_frame.pack(side=tk.LEFT)
        ctk.CTkButton(preview_frame, text="预览选中文件", command=self.app.merge_handler.preview_data,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(preview_frame, text="预览合并结果", command=self.app.merge_handler.preview_merged_data,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        
        # 执行按钮
        ctk.CTkButton(self, text="立即执行合并", command=self.app.merge_handler.merge_files,
                     **self.app.style_config.button_style).pack(side=tk.RIGHT, padx=5)
        
        # 状态栏
        status_frame = ctk.CTkFrame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        ctk.CTkLabel(status_frame, textvariable=self.app.status_var,
                    **self.app.style_config.label_style).pack(fill=tk.X) 