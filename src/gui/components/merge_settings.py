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
        self.entries = {}  # 保存所有输入框的引用
        self.create_widgets()
        self.setup_bindings()
        
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
        self.entries['start_row'] = ctk.CTkEntry(row_frame, textvariable=self.app.merge_config.start_row,
                    width=100, **self.app.style_config.entry_style)
        self.entries['start_row'].pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(row_frame, text="结束行：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        self.entries['end_row'] = ctk.CTkEntry(row_frame, textvariable=self.app.merge_config.end_row,
                    width=100, **self.app.style_config.entry_style)
        self.entries['end_row'].pack(side=tk.LEFT, padx=5)
        
        # 列设置
        col_frame = ctk.CTkFrame(range_frame)
        col_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(col_frame, text="起始列：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        self.entries['start_col'] = ctk.CTkEntry(col_frame, textvariable=self.app.merge_config.start_col,
                    width=100, **self.app.style_config.entry_style)
        self.entries['start_col'].pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(col_frame, text="结束列：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        self.entries['end_col'] = ctk.CTkEntry(col_frame, textvariable=self.app.merge_config.end_col,
                    width=100, **self.app.style_config.entry_style)
        self.entries['end_col'].pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(range_frame, text="注：列请使用Excel列标（如：A、B、C...）", **self.app.style_config.label_style).pack(padx=5, pady=5)
        
        # 表头设置
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        header_settings = ctk.CTkFrame(header_frame)
        header_settings.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(header_settings, text="表头行号：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        self.entries['header_row'] = ctk.CTkEntry(header_settings, textvariable=self.app.merge_config.header_row,
                    width=100, **self.app.style_config.entry_style)
        self.entries['header_row'].pack(side=tk.LEFT, padx=5)
        
        ctk.CTkLabel(header_settings, text="（第几行是表头，从1开始）", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkCheckBox(header_settings, text="保留表头", variable=self.app.merge_config.keep_header,
                       **self.app.style_config.checkbox_style).pack(side=tk.LEFT, padx=20) 
        
    def enable_all_entries(self):
        """启用所有输入框"""
        for entry in self.entries.values():
            entry.configure(state="normal")
            
    def disable_all_entries(self):
        """禁用所有输入框"""
        for entry in self.entries.values():
            entry.configure(state="disabled") 
        
    def setup_bindings(self):
        """设置输入框变更事件绑定"""
        def on_value_change(*args):
            # 如果预览窗口存在，则更新预览
            if (self.app.preview_window.preview_window and 
                self.app.preview_window.preview_window.winfo_exists()):
                try:
                    # 获取当前选中的文件
                    selection = self.app.file_selector.file_tree.selection()
                    if selection:
                        file_path = self.app.file_handler.get_file_path_from_item(selection[0])
                        if file_path and file_path in self.app.file_handler.selected_sheets:
                            # 重新读取数据
                            df = self.app.excel_merger.read_excel_range(
                                file_path,
                                self.app.file_handler.selected_sheets[file_path],
                                self.app.merge_config.header_row.get(),
                                self.app.merge_config.start_row.get(),
                                self.app.merge_config.end_row.get(),
                                self.app.merge_config.start_col.get(),
                                self.app.merge_config.end_col.get()
                            )
                            # 更新预览
                            self.app.preview_window.update_preview(df)
                except Exception as e:
                    print(f"更新预览时出错：{str(e)}")
        
        # 为所有配置变量添加跟踪
        self.app.merge_config.header_row.trace_add("write", on_value_change)
        self.app.merge_config.start_row.trace_add("write", on_value_change)
        self.app.merge_config.end_row.trace_add("write", on_value_change)
        self.app.merge_config.start_col.trace_add("write", on_value_change)
        self.app.merge_config.end_col.trace_add("write", on_value_change) 