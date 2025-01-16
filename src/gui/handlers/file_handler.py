"""
文件处理模块
处理文件选择、清除等操作
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class FileHandler:
    def __init__(self, app):
        """初始化文件处理器"""
        self.app = app
        self.input_files = []
        self.file_sheets = {}  # {文件路径: [sheet名称列表]}
        self.selected_sheets = {}  # {文件路径: 选中的sheet名称}
        
    def add_files(self):
        """添加Excel文件"""
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        for file in files:
            if file not in self.input_files:
                try:
                    # 读取文件的sheet列表
                    xl = pd.ExcelFile(file)
                    sheets = xl.sheet_names
                    
                    # 检查是否为空文件
                    if not sheets:
                        raise ValueError("文件不包含任何Sheet")
                        
                    self.file_sheets[file] = sheets
                    self.selected_sheets[file] = sheets[0]  # 默认选择第一个sheet
                    self.input_files.append(file)
                    
                    # 添加到树形视图
                    file_name = os.path.basename(file)
                    sheet_name = os.path.splitext(file_name)[0] if self.app.merge_config.sheet_name_mode.get() == "auto" else ""
                    self.app.file_selector.file_tree.insert("", tk.END, values=(file_name, sheets[0], sheet_name))
                    
                except Exception as e:
                    messagebox.showerror("错误", f"读取文件 {os.path.basename(file)} 时出错：{str(e)}")
                    continue
                    
        # 更新状态
        self.app.status_var.set(f"已添加 {len(self.input_files)} 个文件")
        
        # 如果是多sheet模式，更新sheet名称
        if self.app.merge_config.merge_mode.get() == "multiple":
            self.on_sheet_name_mode_change()
            
    def clear_files(self):
        """清除所有已添加的文件"""
        self.input_files = []
        self.file_sheets = {}
        self.selected_sheets = {}
        for item in self.app.file_selector.file_tree.get_children():
            self.app.file_selector.file_tree.delete(item)
        self.app.status_var.set("就绪")
        
    def change_sheet(self, item):
        """修改选中文件的Sheet"""
        file_path = self.get_file_path_from_item(item)
        if file_path not in self.file_sheets:
            return
            
        # 创建Sheet选择窗口
        sheet_window = tk.Toplevel(self.app.root)
        sheet_window.title("选择Sheet")
        sheet_window.geometry("300x200")
        sheet_window.transient(self.app.root)
        sheet_window.grab_set()
        
        # 创建Sheet列表
        sheet_list = tk.Listbox(sheet_window, width=40, height=10)
        sheet_list.pack(pady=10)
        
        # 添加Sheet选项
        for sheet in self.file_sheets[file_path]:
            sheet_list.insert(tk.END, sheet)
            
        # 如果已经选择了Sheet，选中它
        if file_path in self.selected_sheets:
            try:
                index = self.file_sheets[file_path].index(self.selected_sheets[file_path])
                sheet_list.selection_set(index)
            except ValueError:
                pass
                
        def confirm_selection():
            selection = sheet_list.curselection()
            if selection:
                selected_sheet = sheet_list.get(selection[0])
                self.selected_sheets[file_path] = selected_sheet
                self.app.file_selector.file_tree.set(item, "选择Sheet", selected_sheet)
                # 如果当前是使用原sheet名模式，更新sheet名称
                if self.app.merge_config.sheet_name_mode.get() == "original":
                    self.app.file_selector.file_tree.set(item, "自定义Sheet名", selected_sheet)
                sheet_window.destroy()
            else:
                messagebox.showwarning("警告", "请先选择一个Sheet！")
                
        def cancel_selection():
            sheet_window.destroy()
        
        # 创建按钮框架
        btn_frame = tk.Frame(sheet_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 添加确定和取消按钮
        tk.Button(btn_frame, text="确定", command=confirm_selection).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="取消", command=cancel_selection).pack(side=tk.RIGHT, padx=5)
        
        # 双击选择功能
        sheet_list.bind('<Double-Button-1>', lambda e: confirm_selection())
        
        # 设置窗口焦点并等待
        sheet_window.focus_set()
        self.app.root.wait_window(sheet_window)
        
    def change_sheet_name(self, item):
        """修改选中文件的Sheet名称"""
        if self.app.merge_config.merge_mode.get() != "multiple" or \
           self.app.merge_config.sheet_name_mode.get() != "custom":
            messagebox.showinfo("提示", "只有在分Sheet保存且使用自定义名称模式下才能修改Sheet名称")
            return
            
        current_name = self.app.file_selector.file_tree.item(item)['values'][2]
        
        # 创建输入对话框
        dialog = tk.Toplevel(self.app.root)
        dialog.title("修改Sheet名称")
        dialog.geometry("300x120")
        
        tk.Label(dialog, text="请输入新的Sheet名称：").pack(padx=10, pady=5)
        entry = tk.Entry(dialog, width=30)
        entry.insert(0, current_name)
        entry.pack(padx=10, pady=5)
        
        def on_ok():
            new_name = entry.get()
            if new_name:
                # 检查新名称是否会造成冲突
                existing_names = []
                for tree_item in self.app.file_selector.file_tree.get_children():
                    if tree_item != item:  # 排除当前项
                        existing_names.append(self.app.file_selector.file_tree.item(tree_item)['values'][2])
                
                if new_name in existing_names:
                    messagebox.showerror("错误", f"Sheet名称 '{new_name}' 已存在，请使用其他名称！")
                    return
                    
                self.app.file_selector.file_tree.set(item, "自定义Sheet名", new_name)
                dialog.destroy()
            
        tk.Button(dialog, text="确定", command=on_ok).pack(pady=10)
        
        # 设置模态对话框
        dialog.transient(self.app.root)
        dialog.grab_set()
        dialog.focus_set()
        
    def get_file_path_from_item(self, item):
        """从树形视图项获取文件路径"""
        file_name = self.app.file_selector.file_tree.item(item)['values'][0]
        for file_path in self.input_files:
            if os.path.basename(file_path) == file_name:
                return file_path
        return None
        
    def on_merge_mode_change(self):
        """当合并模式改变时的处理"""
        if self.app.merge_config.merge_mode.get() == "single":
            self.app.merge_settings.multiple_sheet_frame.pack_forget()
            self.app.merge_settings.single_sheet_frame.pack(fill=tk.X, padx=5, pady=5)
            # 隐藏自定义Sheet名列
            self.app.file_selector.file_tree.column("自定义Sheet名", width=0)
        else:
            self.app.merge_settings.single_sheet_frame.pack_forget()
            self.app.merge_settings.multiple_sheet_frame.pack(fill=tk.X, padx=5, pady=5)
            # 显示自定义Sheet名列
            self.app.file_selector.file_tree.column("自定义Sheet名", width=200)
            self.on_sheet_name_mode_change()
            
    def on_sheet_name_mode_change(self):
        """当Sheet命名方式改变时的处理"""
        if self.app.merge_config.merge_mode.get() == "multiple":
            for item in self.app.file_selector.file_tree.get_children():
                file_name = self.app.file_selector.file_tree.item(item)['values'][0]
                file_path = self.get_file_path_from_item(item)
                current_sheet = self.app.file_selector.file_tree.item(item)['values'][1]
                
                if self.app.merge_config.sheet_name_mode.get() == "auto":
                    # 使用文件名作为sheet名
                    sheet_name = os.path.splitext(file_name)[0]
                    self.app.file_selector.file_tree.set(item, "自定义Sheet名", sheet_name)
                elif self.app.merge_config.sheet_name_mode.get() == "original":
                    # 使用原sheet名
                    self.app.file_selector.file_tree.set(item, "自定义Sheet名", current_sheet)
                else:
                    # 保持当前的自定义名称，如果没有则使用文件名
                    current_name = self.app.file_selector.file_tree.item(item)['values'][2]
                    if not current_name:
                        sheet_name = os.path.splitext(file_name)[0]
                        self.app.file_selector.file_tree.set(item, "自定义Sheet名", sheet_name)
            
            # 检查是否有名称冲突
            conflicts = self.check_sheet_name_conflicts()
            if conflicts:
                conflict_files = [f[1] for f in conflicts]
                messagebox.showwarning("警告", 
                    f"检测到以下文件的Sheet名称存在冲突：\n{', '.join(conflict_files)}\n"
                    "您可以：\n"
                    "1. 使用自定义名称模式手动修改sheet名称\n"
                    "2. 在保存时系统会提供冲突解决方案")
                    
    def check_sheet_name_conflicts(self):
        """检查sheet名称是否有冲突"""
        sheet_names = []
        conflicts = []
        
        for item in self.app.file_selector.file_tree.get_children():
            sheet_name = self.app.file_selector.file_tree.item(item)['values'][2]
            file_name = self.app.file_selector.file_tree.item(item)['values'][0]
            if sheet_name in sheet_names:
                conflicts.append((sheet_name, file_name))
            else:
                sheet_names.append(sheet_name)
                
        return conflicts 