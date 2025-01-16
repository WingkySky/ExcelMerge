"""
GUI样式配置模块
处理界面样式的统一配置
"""
import customtkinter as ctk

class StyleConfig:
    def __init__(self):
        """初始化样式配置"""
        self.configure_styles()
        
    def configure_styles(self):
        """配置所有自定义样式"""
        # 设置全局字体
        self.default_font = ('Microsoft YaHei UI', 12)  # 使用微软雅黑作为默认字体
        
        # 配置按钮样式
        self.button_style = {
            "font": self.default_font,
            "corner_radius": 6,
            "hover": True
        }
        
        # 配置标签样式
        self.label_style = {
            "font": self.default_font,
            "text_color": ("black", "white")  # (light mode, dark mode)
        }
        
        # 配置输入框样式
        self.entry_style = {
            "font": self.default_font,
            "corner_radius": 6
        }
        
        # 配置下拉框样式
        self.combobox_style = {
            "font": self.default_font,
            "corner_radius": 6,
            "button_color": "#2CC985",
            "button_hover_color": "#0C955A"
        }
        
        # 配置单选按钮样式
        self.radio_style = {
            "font": self.default_font,
            "corner_radius": 1000,
            "hover": True,
            "border_width_checked": 6,
            "border_width_unchecked": 3
        }
        
        # 配置复选框样式
        self.checkbox_style = {
            "font": self.default_font,
            "corner_radius": 6,
            "hover": True,
            "border_width": 3
        } 