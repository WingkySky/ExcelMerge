"""
Excel合并工具主程序
"""
import tkinter as tk
import customtkinter as ctk
from src.gui.main_window import ExcelMergerApp

def main():
    # 设置CustomTkinter的外观模式和默认颜色主题
    ctk.set_appearance_mode("light")  # 模式选项: light, dark, system
    ctk.set_default_color_theme("blue")  # 主题选项: blue, dark-blue, green
    
    # 创建主窗口
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 