"""
定时任务设置组件模块
处理定时任务设置界面的创建和交互
"""
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from datetime import datetime

from ...scheduler.task_config import TaskConfig

class ScheduleSettings(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化定时任务设置器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.current_task = None
        self.create_widgets()
        self.refresh_task_list()  # 初始化时刷新任务列表
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="定时任务设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        # 创建左右分栏
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 左侧任务列表
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        task_list_label = ctk.CTkLabel(left_frame, text="任务列表", **self.app.style_config.label_style)
        task_list_label.pack(pady=5)
        
        # 创建任务列表
        list_frame = ctk.CTkFrame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("任务名称", "执行时间", "状态", "上次执行", "下次执行")
        self.task_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        # 设置列
        for col in columns:
            self.task_tree.heading(col, text=col)
            self.task_tree.column(col, width=100)
            
        # 添加滚动条
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.task_tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.configure(yscrollcommand=vsb.set)
        self.task_tree.pack(fill=tk.BOTH, expand=True)
        
        # 添加任务操作按钮
        btn_frame = ctk.CTkFrame(left_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ctk.CTkButton(btn_frame, text="新建任务", command=self.create_task,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(btn_frame, text="删除任务", command=self.delete_task,
                     **self.app.style_config.button_style).pack(side=tk.LEFT, padx=5)
        
        # 右侧任务详情
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        task_detail_label = ctk.CTkLabel(right_frame, text="任务详情", **self.app.style_config.label_style)
        task_detail_label.pack(pady=5)
        
        # 任务基本信息
        info_frame = ctk.CTkFrame(right_frame)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 任务名称
        name_frame = ctk.CTkFrame(info_frame)
        name_frame.pack(fill=tk.X, pady=2)
        ctk.CTkLabel(name_frame, text="任务名称：", width=80,
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        self.name_var = tk.StringVar()
        ctk.CTkEntry(name_frame, textvariable=self.name_var,
                    width=200, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        # 执行时间
        time_frame = ctk.CTkFrame(info_frame)
        time_frame.pack(fill=tk.X, pady=2)
        ctk.CTkLabel(time_frame, text="执行时间：", width=80,
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        self.time_var = tk.StringVar()
        ctk.CTkEntry(time_frame, textvariable=self.time_var,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkLabel(time_frame, text="（24小时制，如：08:30）",
                    **self.app.style_config.label_style).pack(side=tk.LEFT)
        
        # 任务状态
        status_frame = ctk.CTkFrame(info_frame)
        status_frame.pack(fill=tk.X, pady=2)
        self.enabled_var = tk.BooleanVar()
        ctk.CTkCheckBox(status_frame, text="启用任务", variable=self.enabled_var,
                       command=self.toggle_task, **self.app.style_config.checkbox_style).pack(side=tk.LEFT)
        
        # 保存按钮
        save_frame = ctk.CTkFrame(info_frame)
        save_frame.pack(fill=tk.X, pady=5)
        ctk.CTkButton(save_frame, text="保存任务", command=self.save_task,
                     **self.app.style_config.button_style).pack(side=tk.RIGHT)
        
        # 绑定选择事件
        self.task_tree.bind('<<TreeviewSelect>>', self.on_task_selected)
        
        # 添加说明文本
        note_frame = ctk.CTkFrame(self)
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        
        notes = [
            "定时任务说明：",
            "1. 时间格式为24小时制，如：08:30、14:00、23:45",
            "2. 任务将在每天指定时间自动执行合并操作",
            "3. 请确保在启用任务前已正确设置所有合并参数",
            "4. 定时任务运行时请勿关闭软件",
            "5. 可以随时启用/禁用任务"
        ]
        
        for note in notes:
            ctk.CTkLabel(note_frame, text=note, **self.app.style_config.label_style).pack(anchor=tk.W, padx=5, pady=2)
            
    def refresh_task_list(self):
        """刷新任务列表"""
        # 清空列表
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
            
        # 添加任务
        for task in self.app.task_manager.get_all_tasks():
            status = "启用" if task.enabled else "禁用"
            values = (
                task.task_name,
                task.schedule_time,
                status,
                task.last_run or "从未执行",
                task.next_run or "未设置"
            )
            # 确保task_id作为tag被正确设置
            self.task_tree.insert("", tk.END, values=values, tags=(task.task_id,))
            
    def create_task(self):
        """创建新任务"""
        # 检查当前是否有文件被选择
        if not self.app.file_handler.input_files:
            messagebox.showwarning("警告", "请先在文件选择页面添加要处理的文件！")
            return
            
        # 检查是否设置了输出路径
        if not self.app.output_path or not self.app.output_filename_var.get():
            messagebox.showwarning("警告", "请先设置输出路径和文件名！")
            return
            
        # 创建任务配置对话框
        dialog = ctk.CTkToplevel(self)
        dialog.title("创建新任务")
        dialog.geometry("400x300")
        dialog.transient(self)  # 设置为模态窗口
        dialog.grab_set()
        
        # 添加任务配置选项
        ctk.CTkLabel(dialog, text="任务名称：").pack(pady=5)
        name_var = tk.StringVar(value="新建任务")
        name_entry = ctk.CTkEntry(dialog, textvariable=name_var)
        name_entry.pack(pady=5)
        
        ctk.CTkLabel(dialog, text="执行时间：").pack(pady=5)
        time_var = tk.StringVar(value="08:00")
        time_entry = ctk.CTkEntry(dialog, textvariable=time_var)
        time_entry.pack(pady=5)
        
        ctk.CTkLabel(dialog, text="（24小时制，如：08:30）").pack()
        
        enabled_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="创建后立即启用", variable=enabled_var).pack(pady=10)
        
        def confirm_create():
            # 验证时间格式
            try:
                datetime.strptime(time_var.get(), "%H:%M")
            except ValueError:
                messagebox.showerror("错误", "请输入正确的时间格式（HH:MM）")
                return
                
            # 创建新任务
            task = TaskConfig()
            task.task_name = name_var.get()
            task.schedule_time = time_var.get()
            task.enabled = enabled_var.get()
            
            # 设置任务配置
            task.merge_config = self.app.merge_config.get_merge_config()
            task.input_files = [(f, s) for f, s in zip(
                self.app.file_handler.input_files,
                self.app.file_handler.selected_sheets.values()
            )]
            task.output_path = self.app.output_path
            task.output_filename = self.app.output_filename_var.get()
            
            # 如果启用，设置下次执行时间
            if task.enabled:
                task.update_next_run()
            
            # 添加任务
            self.app.task_manager.add_task(task)
            self.refresh_task_list()
            
            # 选中新建的任务
            for item in self.task_tree.get_children():
                if self.task_tree.item(item)['tags'][0] == task.task_id:
                    self.task_tree.selection_set(item)
                    self.on_task_selected(None)  # 触发选中事件
                    break
                    
            dialog.destroy()
            
        # 添加确认和取消按钮
        btn_frame = ctk.CTkFrame(dialog)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        ctk.CTkButton(btn_frame, text="确认", command=confirm_create).pack(side=tk.LEFT, padx=10, expand=True)
        ctk.CTkButton(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10, expand=True)
        
    def delete_task(self):
        """删除任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的任务！")
            return
            
        if messagebox.askyesno("确认", "确定要删除选中的任务吗？\n删除后无法恢复！"):
            try:
                for item in selection:
                    # 获取选中项的tags（第一个tag是task_id）
                    task_id = self.task_tree.item(item)['tags'][0]
                    if task_id:
                        self.app.task_manager.remove_task(task_id)
                        print(f"正在删除任务：{task_id}")  # 添加调试信息
                
                self.refresh_task_list()
                self.clear_task_detail()
                messagebox.showinfo("成功", "任务已删除！")
            except Exception as e:
                print(f"删除任务时出错：{str(e)}")  # 添加错误信息打印
                messagebox.showerror("错误", f"删除任务失败：{str(e)}")
        
    def save_task(self):
        """保存任务"""
        if not self.current_task:
            messagebox.showwarning("警告", "请先选择要编辑的任务！")
            return
            
        # 验证时间格式
        time_str = self.time_var.get()
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            messagebox.showerror("错误", "请输入正确的时间格式（HH:MM）")
            return
            
        # 验证任务名称
        if not self.name_var.get().strip():
            messagebox.showerror("错误", "任务名称不能为空！")
            return
            
        # 更新任务信息
        self.current_task.task_name = self.name_var.get()
        self.current_task.schedule_time = time_str
        old_enabled = self.current_task.enabled
        new_enabled = self.enabled_var.get()
        self.current_task.enabled = new_enabled
        
        # 如果从禁用变为启用，更新下次执行时间
        if not old_enabled and new_enabled:
            self.current_task.update_next_run()
        
        # 更新任务
        self.app.task_manager.update_task(self.current_task)
        self.refresh_task_list()
        messagebox.showinfo("成功", "任务更新成功！")
        
    def toggle_task(self):
        """切换任务状态"""
        if self.current_task:
            old_enabled = self.current_task.enabled
            new_enabled = self.enabled_var.get()
            
            if old_enabled != new_enabled:
                self.current_task.enabled = new_enabled
                if new_enabled:
                    self.current_task.update_next_run()
                else:
                    self.current_task.next_run = None
                
                self.app.task_manager.update_task(self.current_task)
                self.refresh_task_list()
                
                status = "启用" if new_enabled else "禁用"
                messagebox.showinfo("成功", f"任务已{status}！")
            
    def on_task_selected(self, event):
        """当选中任务时"""
        selection = self.task_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        task_id = self.task_tree.item(item)['tags'][0]
        task = self.app.task_manager.get_task(task_id)
        if task:
            self.current_task = task
            self.name_var.set(task.task_name)
            self.time_var.set(task.schedule_time)
            self.enabled_var.set(task.enabled)
            
    def clear_task_detail(self):
        """清空任务详情"""
        self.current_task = None
        self.name_var.set("")
        self.time_var.set("")
        self.enabled_var.set(False) 