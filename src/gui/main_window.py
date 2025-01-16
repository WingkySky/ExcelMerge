"""
主窗口模块
处理主窗口的创建和管理
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from datetime import datetime
import os

from .components.file_selector import FileSelector
from .components.merge_settings import MergeSettings
from .components.excel_style import ExcelStyle
from .components.ui_settings import UISettings
from .components.schedule_settings import ScheduleSettings
from .components.bottom_frame import BottomFrame

from .config.style_config import StyleConfig
from .config.merge_config import MergeConfig

from .handlers.file_handler import FileHandler
from .handlers.merge_handler import MergeHandler
from .handlers.path_manager import PathManager

from .preview.preview_window import PreviewWindow

class ExcelMergerApp:
    def __init__(self, root):
        """初始化主窗口"""
        self.root = root
        self.root.title("Excel文件合并工具")
        
        # 设置窗口大小和位置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 初始化配置
        self.init_config()
        
        # 创建界面
        self.create_gui()
        
    def init_config(self):
        """初始化配置"""
        # 创建样式配置
        self.style_config = StyleConfig()
        
        # 创建合并配置
        self.merge_config = MergeConfig()
        
        # 创建路径管理器
        self.path_manager = PathManager()
        
        # 创建处理器
        self.file_handler = FileHandler(self)
        self.merge_handler = MergeHandler(self)
        
        # 创建预览窗口
        self.preview_window = PreviewWindow(self.root)
        
        # 初始化变量
        self.output_path = os.path.expanduser("~/Documents")  # 默认输出路径
        self.output_path_var = tk.StringVar(value=self.output_path)
        self.output_filename_var = tk.StringVar(value="合并结果")
        self.status_var = tk.StringVar(value="就绪")
        self.time_var = tk.StringVar(value="")
        self.appearance_mode_var = tk.StringVar(value="跟随系统")
        self.color_theme_var = tk.StringVar(value="blue")
        
    def create_gui(self):
        """创建图形界面"""
        # 创建主框架
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建选项卡
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill=tk.BOTH, expand=True)
        
        # 添加选项卡页面
        self.file_page = self.tabview.add("文件选择")
        self.merge_page = self.tabview.add("合并设置")
        self.style_page = self.tabview.add("Excel样式")
        self.ui_page = self.tabview.add("界面设置")
        self.schedule_page = self.tabview.add("定时任务")
        
        # 创建各个页面的内容
        self.file_selector = FileSelector(self.file_page, self)
        self.file_selector.pack(fill=tk.BOTH, expand=True)
        
        self.merge_settings = MergeSettings(self.merge_page, self)
        self.merge_settings.pack(fill=tk.BOTH, expand=True)
        
        self.excel_style = ExcelStyle(self.style_page, self)
        self.excel_style.pack(fill=tk.BOTH, expand=True)
        
        self.ui_settings = UISettings(self.ui_page, self)
        self.ui_settings.pack(fill=tk.BOTH, expand=True)
        
        self.schedule_settings = ScheduleSettings(self.schedule_page, self)
        self.schedule_settings.pack(fill=tk.BOTH, expand=True)
        
        # 创建底部框架
        self.bottom_frame = BottomFrame(self.main_frame, self)
        self.bottom_frame.pack(fill=tk.X, pady=10)
        
    def select_output_path(self):
        """选择输出路径"""
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_path
        )
        if path:
            self.output_path = path
            self.output_path_var.set(path)
            
    def on_path_selected(self, path):
        """当从下拉菜单选择路径时"""
        if path:
            self.output_path = path
            self.output_path_var.set(path)
            
    def change_appearance_mode(self, mode):
        """切换外观模式"""
        mode_map = {
            "跟随系统": "system",
            "浅色": "light",
            "深色": "dark"
        }
        ctk.set_appearance_mode(mode_map.get(mode, "system"))
        
    def change_color_theme(self, theme):
        """切换颜色主题"""
        ctk.set_default_color_theme(theme)
        
    def toggle_schedule(self):
        """切换定时任务状态"""
        if self.schedule_settings.schedule_button.cget("text") == "启动定时任务":
            # 检查时间格式
            time_str = self.time_var.get()
            try:
                datetime.strptime(time_str, "%H:%M")
                self.schedule_settings.schedule_button.configure(text="停止定时任务")
                self.status_var.set(f"定时任务已启动，将在每天 {time_str} 执行")
            except ValueError:
                messagebox.showerror("错误", "请输入正确的时间格式（HH:MM）")
        else:
            self.schedule_settings.schedule_button.configure(text="启动定时任务")
            self.status_var.set("定时任务已停止") 