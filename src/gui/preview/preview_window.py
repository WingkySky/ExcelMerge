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
        self.preview_window = None  # 保存预览窗口实例
        self.current_data = None  # 当前显示的数据
        self.current_tree = None  # 当前的树形视图
        self.on_close_callback = None  # 窗口关闭时的回调函数
        
    def on_window_close(self):
        """处理窗口关闭事件"""
        if self.on_close_callback:
            self.on_close_callback()
        if self.preview_window:
            self.preview_window.destroy()
            self.preview_window = None
            
    def close_existing_preview(self):
        """关闭已存在的预览窗口"""
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.destroy()
            
    def update_preview(self, data):
        """更新预览数据"""
        if not self.preview_window or not self.preview_window.winfo_exists() or not self.current_tree:
            return
            
        # 清空现有数据
        for item in self.current_tree.get_children():
            self.current_tree.delete(item)
            
        # 添加新数据
        for idx, row in data.iterrows():
            if idx < 1000:  # 限制显示前1000行
                values = [str(row[col]) for col in data.columns]
                self.current_tree.insert("", tk.END, values=values)
            else:
                break
                
        # 更新统计信息
        total_rows = len(data)
        shown_rows = min(total_rows, 1000)
        self.update_stats(total_rows, shown_rows)
        
    def update_stats(self, total_rows, shown_rows):
        """更新统计信息"""
        if not self.preview_window or not self.preview_window.winfo_exists():
            return
            
        # 更新统计标签
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        ttk.Label(self.info_frame, text=f"总行数: {total_rows}    显示行数: {shown_rows}").pack(side=tk.LEFT)
        
        if total_rows > 1000:
            ttk.Label(self.info_frame, text="（仅显示前1000行）", foreground="red").pack(side=tk.LEFT, padx=5)
            
    def show_preview(self, df, title="预览数据", on_close=None):
        """
        显示数据预览窗口
        Args:
            df: 要显示的DataFrame
            title: 窗口标题
            on_close: 窗口关闭时的回调函数
        """
        # 保存回调函数
        self.on_close_callback = on_close
        
        # 关闭已存在的预览窗口
        self.close_existing_preview()
        
        # 创建新的预览窗口
        self.preview_window = tk.Toplevel(self.parent)
        self.preview_window.title(title)
        
        # 设置窗口关闭事件
        self.preview_window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        # 设置窗口大小
        screen_width = self.preview_window.winfo_screenwidth()
        screen_height = self.preview_window.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 创建主框架
        main_frame = ttk.Frame(self.preview_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建树形视图
        tree = ttk.Treeview(main_frame)
        
        # 创建垂直滚动条
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建水平滚动条
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 配置树形视图的滚动
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 设置列
        columns = list(df.columns)
        tree["columns"] = columns
        tree["show"] = "headings"  # 不显示第一个空列
        
        # 设置每列的标题和宽度
        for col in columns:
            tree.heading(col, text=str(col))
            # 计算列宽度（根据数据内容）
            max_width = max(
                len(str(col)),  # 标题长度
                df[col].astype(str).str.len().max() if len(df) > 0 else 0  # 数据长度
            )
            tree.column(col, width=min(max_width * 10, 300))  # 限制最大宽度为300像素
        
        # 添加数据行
        for idx, row in df.iterrows():
            if idx < 1000:  # 限制显示前1000行
                values = [str(row[col]) for col in columns]
                tree.insert("", tk.END, values=values)
            else:
                break
        
        # 添加数据统计信息
        info_frame = ttk.Frame(self.preview_window)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 显示总行数和当前显示行数
        total_rows = len(df)
        shown_rows = min(total_rows, 1000)
        ttk.Label(info_frame, text=f"总行数: {total_rows}    显示行数: {shown_rows}").pack(side=tk.LEFT)
        
        # 如果数据被截断，显示提示信息
        if total_rows > 1000:
            ttk.Label(info_frame, text="（仅显示前1000行）", foreground="red").pack(side=tk.LEFT, padx=5)
        
        # 添加关闭按钮
        ttk.Button(self.preview_window, text="关闭", command=self.preview_window.destroy).pack(pady=5)
        
    def show_multi_sheet_preview(self, all_data, title="预览: 多Sheet结果", on_close=None):
        """
        显示多Sheet数据预览窗口
        Args:
            all_data: 要显示的数据列表
            title: 窗口标题
            on_close: 窗口关闭时的回调函数
        """
        # 保存回调函数
        self.on_close_callback = on_close
        
        # 关闭已存在的预览窗口
        self.close_existing_preview()
        
        # 创建新的预览窗口
        self.preview_window = tk.Toplevel(self.parent)
        self.preview_window.title(title)
        
        # 设置窗口关闭事件
        self.preview_window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        # 设置窗口大小
        screen_width = self.preview_window.winfo_screenwidth()
        screen_height = self.preview_window.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 创建Notebook
        notebook = ttk.Notebook(self.preview_window)
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
        ttk.Button(self.preview_window, text="关闭", command=self.on_window_close).pack(pady=5) 