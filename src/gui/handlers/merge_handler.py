"""
合并处理模块
处理Excel文件合并相关的操作
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

class MergeHandler:
    def __init__(self, app):
        """初始化合并处理器"""
        self.app = app
        
    def merge_files(self):
        """执行Excel文件合并操作"""
        if not self.app.file_handler.input_files:
            messagebox.showerror("错误", "请先选择要合并的Excel文件！")
            return
            
        if not self.app.output_path:
            messagebox.showerror("错误", "请选择输出路径！")
            return
            
        try:
            # 准备合并参数
            merge_config = self.app.merge_config.get_merge_config()
            
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.app.output_filename_var.get()}_{timestamp}.xlsx"
            output_file = os.path.join(self.app.output_path, filename)
            
            # 让用户确认或修改文件名
            output_file = filedialog.asksaveasfilename(
                initialdir=self.app.output_path,
                initialfile=filename,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="保存合并文件"
            )
            
            if not output_file:  # 用户取消了保存
                return
                
            # 更新输出路径
            self.app.output_path = os.path.dirname(output_file)
            self.app.output_path_var.set(self.app.output_path)
            
            # 添加到最近使用路径
            self.app.path_manager.add_recent_path(self.app.output_path)
            
            # 执行合并操作
            result = self.app.excel_merger.merge_files(
                self.app.file_handler.input_files,
                output_file,
                self.app.file_handler.selected_sheets,
                self.app.file_handler.file_sheets,
                merge_config
            )
            
            if result['success']:
                self.app.status_var.set(f"合并完成！输出文件：{output_file}")
                messagebox.showinfo("成功", f"文件合并完成！\n共合并了 {len(self.app.file_handler.input_files)} 个文件的数据")
            else:
                raise Exception(result['error'])
                
        except Exception as e:
            self.app.status_var.set(f"错误：{str(e)}")
            messagebox.showerror("错误", f"合并过程中出现错误：{str(e)}")
            
    def preview_data(self):
        """预览选中的文件"""
        selection = self.app.file_selector.file_tree.selection()
        if not selection:
            messagebox.showerror("错误", "请先选择要预览的文件！")
            return
            
        file_path = self.app.file_handler.get_file_path_from_item(selection[0])
        if not file_path or file_path not in self.app.file_handler.selected_sheets:
            messagebox.showerror("错误", "请先为文件选择Sheet！")
            return
            
        try:
            df = self.app.excel_merger.read_excel_range(
                file_path,
                self.app.file_handler.selected_sheets[file_path],
                self.app.merge_config.header_row.get(),
                self.app.merge_config.start_row.get(),
                self.app.merge_config.end_row.get(),
                self.app.merge_config.start_col.get(),
                self.app.merge_config.end_col.get()
            )
            
            # 显示预览窗口
            self.app.preview_window.show_preview(df, f"预览: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"预览数据时出错：{str(e)}")
            
    def preview_merged_data(self):
        """预览合并后的数据"""
        if not self.app.file_handler.input_files:
            messagebox.showerror("错误", "请先选择要合并的Excel文件！")
            return
            
        try:
            # 准备合并参数
            merge_config = self.app.merge_config.get_merge_config()
            
            # 读取所有Excel文件的指定范围
            all_data = []
            for file in self.app.file_handler.input_files:
                if file in self.app.file_handler.selected_sheets:
                    df = self.app.excel_merger.read_excel_range(
                        file,
                        self.app.file_handler.selected_sheets[file],
                        merge_config['header_row'],
                        merge_config['start_row'],
                        merge_config['end_row'],
                        merge_config['start_col'],
                        merge_config['end_col'],
                        add_source=(merge_config['merge_mode'] == 'single')
                    )
                    if not df.empty:
                        all_data.append((file, df))
                        
            if not all_data:
                raise ValueError("没有有效的数据可以合并！")
                
            if merge_config['merge_mode'] == "single":
                # 检查表头一致性
                headers_consistent, message = self.app.excel_merger.check_headers_consistency(
                    [df for _, df in all_data]
                )
                if not headers_consistent:
                    if not messagebox.askyesno("警告", f"发现表头不一致：\n{message}\n是否继续预览？"):
                        return
                    
                # 合并数据
                merged_df = self.app.excel_merger.smart_merge(
                    [df for _, df in all_data],
                    merge_config['keep_header']
                )
                self.app.preview_window.show_preview(merged_df, "预览: 合并结果")
            else:
                # 准备多sheet预览数据
                preview_data = []
                for file_path, df in all_data:
                    # 获取sheet名称
                    file_name = os.path.basename(file_path)
                    sheet_name = None
                    for item in self.app.file_selector.file_tree.get_children():
                        if self.app.file_selector.file_tree.item(item)['values'][0] == file_name:
                            sheet_name = self.app.file_selector.file_tree.item(item)['values'][2]
                            break
                    
                    if not sheet_name:
                        sheet_name = os.path.splitext(file_name)[0]
                        
                    preview_data.append((sheet_name, df))
                    
                self.app.preview_window.show_multi_sheet_preview(preview_data)
                
        except Exception as e:
            messagebox.showerror("错误", f"预览数据时出错：{str(e)}") 