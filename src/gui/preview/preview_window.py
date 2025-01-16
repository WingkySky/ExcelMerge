"""
预览窗口模块
处理数据预览相关的功能
"""
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

class PreviewWindow:
    def __init__(self, parent):
        """初始化预览窗口"""
        self.parent = parent
        
    def show_preview(self, data, title):
        """显示单个数据预览窗口"""
        preview_window = tk.Toplevel(self.parent)
        preview_window.title(title)
        preview_window.geometry("1000x600")
        
        # 创建主框架
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建树形视图
        tree = ttk.Treeview(main_frame)
        
        # 创建滚动条
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 配置树形视图的滚动
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 设置列
        columns = list(data.columns)
        tree["columns"] = columns
        tree["show"] = "headings"
        
        # 设置列标题和宽度
        for col in columns:
            tree.heading(col, text=str(col))
            max_width = max(
                len(str(col)),
                data[col].astype(str).str.len().max() if len(data) > 0 else 0
            )
            tree.column(col, width=min(max_width * 10, 300))
        
        # 添加数据行
        for idx, row in data.iterrows():
            if idx < 1000:  # 限制显示前1000行
                values = [str(row[col]) for col in columns]
                tree.insert("", tk.END, values=values)
            else:
                break
        
        # 添加数据统计信息
        info_frame = ttk.Frame(preview_window)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        total_rows = len(data)
        shown_rows = min(total_rows, 1000)
        ttk.Label(info_frame, text=f"总行数: {total_rows}    显示行数: {shown_rows}").pack(side=tk.LEFT)
        
        if total_rows > 1000:
            ttk.Label(info_frame, text="（仅显示前1000行）", foreground="red").pack(side=tk.LEFT, padx=5)
        
        # 添加关闭按钮
        ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=5)
        
    def show_multi_sheet_preview(self, all_data, title="预览: 多Sheet结果"):
        """显示多Sheet数据预览窗口"""
        preview_window = tk.Toplevel(self.parent)
        preview_window.title(title)
        
        # 设置窗口大小
        screen_width = preview_window.winfo_screenwidth()
        screen_height = preview_window.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 创建Notebook
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 为每个文件创建一个sheet页
        for sheet_name, df in all_data:
            # 创建sheet页
            sheet_frame = ttk.Frame(notebook)
            notebook.add(sheet_frame, text=sheet_name)
            
            # 创建表格
            tree = ttk.Treeview(sheet_frame)
            
            # 创建滚动条
            vsb = ttk.Scrollbar(sheet_frame, orient="vertical", command=tree.yview)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            
            hsb = ttk.Scrollbar(sheet_frame, orient="horizontal", command=tree.xview)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            
            # 配置树形视图的滚动
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # 设置列
            columns = list(df.columns)
            tree["columns"] = columns
            tree["show"] = "headings"
            
            # 设置列标题和宽度
            for col in columns:
                tree.heading(col, text=str(col))
                max_width = max(
                    len(str(col)),
                    df[col].astype(str).str.len().max() if len(df) > 0 else 0
                )
                tree.column(col, width=min(max_width * 10, 300))
            
            # 添加数据行
            for idx, row in df.iterrows():
                if idx < 1000:
                    values = [str(row[col]) for col in columns]
                    tree.insert("", tk.END, values=values)
                else:
                    break
            
            # 添加数据统计信息
            info_frame = ttk.Frame(sheet_frame)
            info_frame.pack(fill=tk.X, padx=10, pady=5)
            
            total_rows = len(df)
            shown_rows = min(total_rows, 1000)
            ttk.Label(info_frame, text=f"总行数: {total_rows}    显示行数: {shown_rows}").pack(side=tk.LEFT)
            
            if total_rows > 1000:
                ttk.Label(info_frame, text="（仅显示前1000行）", foreground="red").pack(side=tk.LEFT, padx=5)
        
        # 添加关闭按钮
        ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=5) 