"""
GUI样式配置模块
包含所有CustomTkinter的样式配置
"""
import customtkinter as ctk

class StyleConfig:
    def __init__(self):
        """初始化样式配置"""
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

    def set_theme(self, mode=None, color_theme=None):
        """设置主题和颜色"""
        if mode:
            mode_map = {
                "跟随系统": "system",
                "浅色": "light",
                "深色": "dark"
            }
            ctk.set_appearance_mode(mode_map.get(mode, mode))
            
        if color_theme:
            ctk.set_default_color_theme(color_theme) 