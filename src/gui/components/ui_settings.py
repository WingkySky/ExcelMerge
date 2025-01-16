"""
界面设置组件模块
处理界面设置界面的创建和交互
"""
import tkinter as tk
import customtkinter as ctk

class UISettings(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化界面设置器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="软件界面设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        # 添加外观模式选择
        theme_frame = ctk.CTkFrame(self)
        theme_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(theme_frame, text="外观模式：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        
        # 创建外观模式选择
        appearance_modes = ["跟随系统", "浅色", "深色"]
        self.appearance_mode_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=appearance_modes,
            command=self.app.change_appearance_mode,
            variable=self.app.appearance_mode_var,
            width=200,
            **self.app.style_config.combobox_style
        )
        self.appearance_mode_menu.pack(side=tk.LEFT, padx=5)
        
        # 添加颜色主题选择
        color_theme_frame = ctk.CTkFrame(self)
        color_theme_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(color_theme_frame, text="颜色主题：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        
        # 创建颜色主题选择
        color_themes = ["blue", "dark-blue", "green"]
        self.color_theme_menu = ctk.CTkOptionMenu(
            color_theme_frame,
            values=color_themes,
            command=self.app.change_color_theme,
            variable=self.app.color_theme_var,
            width=200,
            **self.app.style_config.combobox_style
        )
        self.color_theme_menu.pack(side=tk.LEFT, padx=5)
        
        # 添加说明文本
        note_frame = ctk.CTkFrame(self)
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        
        notes = [
            "界面设置说明：",
            "1. 外观模式：可选择跟随系统、浅色或深色模式",
            "2. 颜色主题：可选择不同的主题色调",
            "注：部分设置可能需要重启软件后生效"
        ]
        
        for note in notes:
            ctk.CTkLabel(note_frame, text=note, **self.app.style_config.label_style).pack(anchor=tk.W, padx=5, pady=2) 