"""
程序入口模块
"""
import os
import sys
import platform

# 在 macOS 上重定向系统日志
if platform.system() == 'Darwin':  # Darwin 是 macOS 的系统名
    # 将标准错误输出重定向到 /dev/null
    sys.stderr = open(os.devnull, 'w')

import tkinter as tk
import customtkinter as ctk
from src.gui.main_window import ExcelMergerApp

def main():
    """主函数"""
    # 设置默认主题
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    
    # 创建主窗口
    root = ctk.CTk()
    app = ExcelMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 