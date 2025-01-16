"""
定时任务设置组件模块
处理定时任务设置界面的创建和交互
"""
import tkinter as tk
import customtkinter as ctk

class ScheduleSettings(ctk.CTkFrame):
    def __init__(self, parent, app, **kwargs):
        """初始化定时任务设置器"""
        super().__init__(parent, **kwargs)
        self.app = app
        self.create_widgets()
        
    def create_widgets(self):
        """创建组件"""
        # 添加标题
        self.title = ctk.CTkLabel(self, text="定时任务设置", font=("Microsoft YaHei UI", 16, "bold"))
        self.title.pack(pady=10)
        
        schedule_settings_frame = ctk.CTkFrame(self)
        schedule_settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 时间设置
        schedule_time_frame = ctk.CTkFrame(schedule_settings_frame)
        schedule_time_frame.pack(fill=tk.X, padx=5, pady=5)
        ctk.CTkLabel(schedule_time_frame, text="设置定时执行时间（24小时制）：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(schedule_time_frame, textvariable=self.app.time_var,
                    width=100, **self.app.style_config.entry_style).pack(side=tk.LEFT, padx=5)
        
        # 启动/停止按钮
        self.schedule_button = ctk.CTkButton(schedule_time_frame, text="启动定时任务", command=self.app.toggle_schedule,
                     **self.app.style_config.button_style)
        self.schedule_button.pack(side=tk.LEFT, padx=20)
        
        # 添加说明文本
        note_frame = ctk.CTkFrame(self)
        note_frame.pack(fill=tk.X, padx=10, pady=10)
        
        notes = [
            "定时任务说明：",
            "1. 时间格式为24小时制，如：08:30、14:00、23:45",
            "2. 定时任务将在每天指定时间自动执行合并操作",
            "3. 请确保在启动定时任务前已正确设置所有合并参数",
            "4. 定时任务运行时请勿关闭软件",
            "5. 可以随时停止定时任务"
        ]
        
        for note in notes:
            ctk.CTkLabel(note_frame, text=note, **self.app.style_config.label_style).pack(anchor=tk.W, padx=5, pady=2)
            
        # 任务状态显示
        status_frame = ctk.CTkFrame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ctk.CTkLabel(status_frame, text="当前状态：", **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5)
        ctk.CTkLabel(status_frame, textvariable=self.app.status_var, **self.app.style_config.label_style).pack(side=tk.LEFT, padx=5) 