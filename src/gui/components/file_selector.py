"""
文件选择组件模块
处理文件选择界面的创建和交互
"""
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

class FileSelector(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化文件选择器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="文件选择与输出设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        # 文件和Sheet选择区域
        file_section_frame = ctk.CTkFrame(self)
        file_section_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建带滚动条的框架
        tree_frame = ctk.CTkFrame(file_section_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ("文件名", "选择Sheet", "自定义Sheet名")
        self.file_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8)
        
        # 添加垂直滚动条
        vsb = ctk.CTkScrollbar(tree_frame, orientation="vertical", command=self.file_tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 添加水平滚动条
        hsb = ctk.CTkScrollbar(tree_frame, orientation="horizontal", command=self.file_tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 配置树形视图的滚动
        self.file_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 设置列标题和宽度
        for col in columns:
            self.file_tree.heading(col, text=col)
            self.file_tree.column(col, width=150)
        
        # 添加和清除按钮
        btn_frame = ctk.CTkFrame(file_section_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ctk.CTkButton(btn_frame, text="添加文件", command=self.app.file_handler.add_files, 
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(btn_frame, text="清除所有", command=self.app.file_handler.clear_files,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(btn_frame, text="修改选中文件的Sheet", 
                     command=lambda: self.app.file_handler.change_sheet(self.file_tree.selection()[0]) if self.file_tree.selection() else None,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(btn_frame, text="修改Sheet名称",
                     command=lambda: self.app.file_handler.change_sheet_name(self.file_tree.selection()[0]) if self.file_tree.selection() else None,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        
        # 输出路径选择
        output_frame = ctk.CTkFrame(self)
        output_frame.pack(fill=tk.X, padx=10, pady=5)
        
        output_label = ctk.CTkLabel(output_frame, text="输出设置", **self.app.style_config.label_style)
        output_label.pack(pady=10)
        
        # 默认路径选择
        path_frame = ctk.CTkFrame(output_frame)
        path_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ctk.CTkLabel(path_frame, text="快速路径：", width=80, 
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        
        # 创建下拉菜单
        self.path_combobox = ctk.CTkComboBox(path_frame, values=self.app.path_manager.get_available_paths(),
                                            width=400, command=self.app.on_path_selected,
                                            **self.app.style_config.combobox_style)
        if self.app.path_manager.get_available_paths():
            self.path_combobox.set(self.app.path_manager.get_available_paths()[0])
        self.path_combobox.pack(side=tk.LEFT, padx=5)
        
        # 自定义路径
        custom_path_frame = ctk.CTkFrame(output_frame)
        custom_path_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ctk.CTkLabel(custom_path_frame, text="自定义路径：", width=80,
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        
        ctk.CTkEntry(custom_path_frame, textvariable=self.app.output_path_var,
                    width=400, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        ctk.CTkButton(custom_path_frame, text="浏览", command=self.app.select_output_path,
                     **self.app.style_config.button_style).pack(side=tk.LEFT)
        
        # 文件名设置
        filename_frame = ctk.CTkFrame(output_frame)
        filename_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ctk.CTkLabel(filename_frame, text="文件名：", width=80,
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        
        ctk.CTkEntry(filename_frame, textvariable=self.app.output_filename_var,
                    width=300, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(filename_frame, text=".xlsx",
                    **self.app.style_config.label_style).pack(side=tk.LEFT) 