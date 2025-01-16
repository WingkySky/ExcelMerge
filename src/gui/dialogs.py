"""
对话框管理模块
处理所有自定义对话框的显示和交互
"""
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

class SheetSelectionDialog:
    def __init__(self, parent, file_path, available_sheets, current_sheet=None, style_config=None):
        """Sheet选择对话框"""
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("选择Sheet")
        self.dialog.geometry("300x200")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.file_path = file_path
        self.available_sheets = available_sheets
        self.current_sheet = current_sheet
        self.style_config = style_config or {}
        self.result = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # 创建Sheet列表
        self.sheet_list = tk.Listbox(self.dialog, width=40, height=10)
        self.sheet_list.pack(pady=10)
        
        # 添加Sheet选项
        for sheet in self.available_sheets:
            self.sheet_list.insert(tk.END, sheet)
            
        # 如果有当前选中的Sheet，选中它
        if self.current_sheet in self.available_sheets:
            try:
                index = self.available_sheets.index(self.current_sheet)
                self.sheet_list.selection_set(index)
            except ValueError:
                pass
                
        # 创建按钮框架
        btn_frame = ctk.CTkFrame(self.dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 添加确定和取消按钮
        ctk.CTkButton(btn_frame, text="确定", command=self.confirm_selection,
                     **self.style_config.get("button_style", {})).pack(side=tk.RIGHT, padx=5)
        ctk.CTkButton(btn_frame, text="取消", command=self.cancel_selection,
                     **self.style_config.get("button_style", {})).pack(side=tk.RIGHT, padx=5)
        
        # 双击选择功能
        self.sheet_list.bind('<Double-Button-1>', lambda e: self.confirm_selection())
        
    def confirm_selection(self):
        selection = self.sheet_list.curselection()
        if selection:
            self.result = self.sheet_list.get(selection[0])
            self.dialog.destroy()
        else:
            messagebox.showwarning("警告", "请先选择一个Sheet！")
            
    def cancel_selection(self):
        self.dialog.destroy()
        
    def show(self):
        self.dialog.focus_set()
        self.dialog.wait_window()
        return self.result

class SheetNameDialog:
    def __init__(self, parent, current_name="", existing_names=None, style_config=None):
        """Sheet名称编辑对话框"""
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("修改Sheet名称")
        self.dialog.geometry("300x120")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.current_name = current_name
        self.existing_names = existing_names or []
        self.style_config = style_config or {}
        self.result = None
        
        self.create_widgets()
        
    def create_widgets(self):
        ctk.CTkLabel(self.dialog, text="请输入新的Sheet名称：",
                    **self.style_config.get("label_style", {})).pack(padx=10, pady=5)
        
        self.entry = ctk.CTkEntry(self.dialog, width=200,
                                **self.style_config.get("entry_style", {}))
        self.entry.insert(0, self.current_name)
        self.entry.pack(padx=10, pady=5)
        
        ctk.CTkButton(self.dialog, text="确定", command=self.confirm,
                     **self.style_config.get("button_style", {})).pack(pady=10)
        
    def confirm(self):
        new_name = self.entry.get()
        if new_name:
            if new_name in self.existing_names:
                messagebox.showerror("错误", f"Sheet名称 '{new_name}' 已存在，请使用其他名称！")
                return
            self.result = new_name
            self.dialog.destroy()
            
    def show(self):
        self.dialog.focus_set()
        self.dialog.wait_window()
        return self.result

class ConflictResolutionDialog:
    def __init__(self, parent, conflicts, existing_names=None, style_config=None):
        """Sheet名称冲突解决对话框"""
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Sheet名称冲突")
        self.dialog.geometry("600x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.conflicts = conflicts
        self.existing_names = existing_names or []
        self.style_config = style_config or {}
        self.entries = {}
        self.result = {"confirmed": False, "names": {}}
        
        self.create_widgets()
        
    def create_widgets(self):
        # 说明文本
        ctk.CTkLabel(self.dialog, text="检测到以下Sheet名称冲突，请选择处理方式：",
                    **self.style_config.get("label_style", {})).pack(padx=10, pady=5)
        
        # 创建滚动框架
        frame = ctk.CTkFrame(self.dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(frame)
        scrollbar = ctk.CTkScrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ctk.CTkFrame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 为每个冲突创建处理选项
        for sheet_name, file_name in self.conflicts:
            conflict_frame = ctk.CTkFrame(scrollable_frame)
            conflict_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ctk.CTkLabel(conflict_frame, text=f"当前Sheet名称: {sheet_name}",
                        **self.style_config.get("label_style", {})).pack(padx=5, pady=2)
            
            # 建议的新名称
            suggested_name = self.suggest_name(sheet_name)
            
            name_frame = ctk.CTkFrame(conflict_frame)
            name_frame.pack(fill=tk.X, padx=5, pady=2)
            
            ctk.CTkLabel(name_frame, text="新名称：",
                        **self.style_config.get("label_style", {})).pack(side=tk.LEFT)
            
            entry = ctk.CTkEntry(name_frame, width=200,
                               **self.style_config.get("entry_style", {}))
            entry.insert(0, suggested_name)
            entry.pack(side=tk.LEFT, padx=5)
            
            # 使用建议按钮
            ctk.CTkButton(name_frame, text="使用建议名称",
                         command=lambda e=entry, s=suggested_name: self.use_suggestion(e, s),
                         **self.style_config.get("button_style", {})).pack(side=tk.LEFT, padx=5)
            
            self.entries[file_name] = entry
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 按钮区域
        btn_frame = ctk.CTkFrame(self.dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ctk.CTkButton(btn_frame, text="确定", command=self.confirm,
                     **self.style_config.get("button_style", {})).pack(side=tk.RIGHT, padx=5)
        ctk.CTkButton(btn_frame, text="取消", command=self.cancel,
                     **self.style_config.get("button_style", {})).pack(side=tk.RIGHT, padx=5)
        
    def suggest_name(self, base_name):
        """为冲突的sheet名称生成建议名称"""
        counter = 1
        new_name = base_name
        while new_name in self.existing_names:
            new_name = f"{base_name}_{counter}"
            counter += 1
        return new_name
        
    def use_suggestion(self, entry, suggested_name):
        """使用建议的名称"""
        entry.delete(0, tk.END)
        entry.insert(0, suggested_name)
        
    def confirm(self):
        """确认修改"""
        # 检查新名称是否仍有冲突
        new_names = [entry.get() for entry in self.entries.values()]
        if len(new_names) != len(set(new_names)):
            messagebox.showerror("错误", "新的Sheet名称仍然存在冲突，请修改后重试！")
            return
            
        self.result["confirmed"] = True
        self.result["names"] = {file_name: entry.get() for file_name, entry in self.entries.items()}
        self.dialog.destroy()
        
    def cancel(self):
        """取消修改"""
        self.dialog.destroy()
        
    def show(self):
        """显示对话框并返回结果"""
        self.dialog.focus_set()
        self.dialog.wait_window()
        return self.result 