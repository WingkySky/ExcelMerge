"""
合并设置组件模块
处理合并设置界面的创建和交互
"""
import tkinter as tk
import customtkinter as ctk

class MergeSettings(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化合并设置器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="合并方式设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        # 合并方式设置
        merge_settings_frame = ctk.CTkFrame(self)
        merge_settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 合并方式选择
        merge_mode_frame = ctk.CTkFrame(merge_settings_frame)
        merge_mode_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(merge_mode_frame, text="合并方式：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkRadioButton(merge_mode_frame, text="合并到单个Sheet", variable=self.app.merge_config.merge_mode, 
                       value="single", command=self.app.file_handler.on_merge_mode_change,
                       **self.app.style_config.radio_style).pack(side=tk.LEFT, padx=20)
        ctk.CTkRadioButton(merge_mode_frame, text="每个文件单独一个Sheet", variable=self.app.merge_config.merge_mode, 
                       value="multiple", command=self.app.file_handler.on_merge_mode_change,
                       **self.app.style_config.radio_style).pack(side=tk.LEFT, padx=20)
        
        # Sheet名称设置（只在单sheet模式下显示）
        self.single_sheet_frame = ctk.CTkFrame(merge_settings_frame)
        ctk.CTkLabel(self.single_sheet_frame, text="Sheet名称：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(self.single_sheet_frame, textvariable=self.app.merge_config.custom_sheet_name,
                    width=300, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        # 多sheet模式下的设置
        self.multiple_sheet_frame = ctk.CTkFrame(merge_settings_frame)
        ctk.CTkLabel(self.multiple_sheet_frame, text="Sheet命名方式：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkRadioButton(self.multiple_sheet_frame, text="使用文件名", variable=self.app.merge_config.sheet_name_mode, 
                       value="auto", command=self.app.file_handler.on_sheet_name_mode_change,
                       **self.app.style_config.radio_style).pack(side=tk.LEFT, padx=20)
        ctk.CTkRadioButton(self.multiple_sheet_frame, text="使用原Sheet名", variable=self.app.merge_config.sheet_name_mode, 
                       value="original", command=self.app.file_handler.on_sheet_name_mode_change,
                       **self.app.style_config.radio_style).pack(side=tk.LEFT, padx=20)
        ctk.CTkRadioButton(self.multiple_sheet_frame, text="使用自定义名称", variable=self.app.merge_config.sheet_name_mode, 
                       value="custom", command=self.app.file_handler.on_sheet_name_mode_change,
                       **self.app.style_config.radio_style).pack(side=tk.LEFT, padx=20)
        
        # 数据区间选择
        range_frame = ctk.CTkFrame(self)
        range_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 行设置
        row_frame = ctk.CTkFrame(range_frame)
        row_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(row_frame, text="起始行：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(row_frame, textvariable=self.app.merge_config.start_row,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkLabel(row_frame, text="结束行：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(row_frame, textvariable=self.app.merge_config.end_row,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        # 列设置
        col_frame = ctk.CTkFrame(range_frame)
        col_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(col_frame, text="起始列：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(col_frame, textvariable=self.app.merge_config.start_col,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkLabel(col_frame, text="结束列：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(col_frame, textvariable=self.app.merge_config.end_col,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(range_frame, text="注：列请使用Excel列标（如：A、B、C...）", **self.app.style_config.label_style).pack(padx=5, pady=5)
        
        # 表头设置
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        header_settings = ctk.CTkFrame(header_frame)
        header_settings.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(header_settings, text="表头行号：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(header_settings, textvariable=self.app.merge_config.header_row,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkLabel(header_settings, text="（第几行是表头，从1开始）", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkCheckBox(header_settings, text="保留表头", variable=self.app.merge_config.keep_header,
                       **self.app.style_config.checkbox_style).pack(side=tk.LEFT, padx=20) 