"""
Excel合并核心逻辑模块
处理Excel文件的读取、合并等操作
"""
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class ExcelMerger:
    def __init__(self, style_manager=None):
        """初始化Excel合并器"""
        self.style_manager = style_manager

    def merge_files(self, input_files, output_file, selected_sheets, file_sheets, merge_config):
        """
        执行Excel文件合并操作
        
        Args:
            input_files: 输入文件列表
            output_file: 输出文件路径
            selected_sheets: 选中的sheet信息 {文件路径: sheet名称}
            file_sheets: 文件的sheet信息 {文件路径: [sheet名称列表]}
            merge_config: 合并配置参数
            
        Returns:
            dict: 包含操作结果的字典
                - success: 是否成功
                - error: 错误信息（如果失败）
        """
        try:
            # 读取所有Excel文件的指定范围
            all_data = []
            first_file = True
            header_styles = None
            data_styles = None
            
            for file in input_files:
                if file in selected_sheets:
                    df = self.read_excel_range(
                        file,
                        selected_sheets[file],
                        merge_config['header_row'],
                        merge_config['start_row'],
                        merge_config['end_row'],
                        merge_config['start_col'],
                        merge_config['end_col'],
                        add_source=(merge_config['merge_mode'] == 'single')
                    )
                    
                    if not df.empty:
                        all_data.append((file, df))
                        
                        # 从第一个文件获取样式模板
                        if first_file and merge_config['keep_styles'] and self.style_manager:
                            try:
                                wb = load_workbook(file)
                                header_styles, data_styles = self.style_manager.get_column_styles(
                                    wb, 
                                    selected_sheets[file],
                                    merge_config['header_row']
                                )
                                first_file = False
                            except Exception as style_error:
                                print(f"获取样式时出错: {style_error}")
                                first_file = False
            
            if not all_data:
                return {'success': False, 'error': "没有有效的数据可以合并！"}
                
            # 根据合并方式处理数据
            if merge_config['merge_mode'] == "single":
                # 合并到单个sheet
                # 检查表头一致性
                headers_consistent, message = self.check_headers_consistency([df for _, df in all_data])
                if not headers_consistent:
                    return {'success': False, 'error': f"表头不一致：\n{message}"}
                
                # 智能合并数据
                merged_df = self.smart_merge([df for _, df in all_data], merge_config['keep_header'])
                
                # 确定sheet名称
                sheet_name = merge_config['custom_sheet_name'] if merge_config['sheet_name_mode'] == "custom" else "合并结果"
                
                # 保存合并后的文件
                self.save_merged_file(
                    output_file,
                    merged_df,
                    sheet_name,
                    header_styles if merge_config['keep_styles'] else None,
                    data_styles if merge_config['keep_styles'] else None,
                    merge_config
                )
            else:
                # 每个文件一个sheet
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for file_path, df in all_data:
                        # 获取sheet名称
                        file_name = os.path.basename(file_path)
                        if merge_config['sheet_name_mode'] == "auto":
                            sheet_name = os.path.splitext(file_name)[0]
                        elif merge_config['sheet_name_mode'] == "original":
                            sheet_name = selected_sheets[file_path]
                        else:  # custom
                            sheet_name = file_sheets[file_path].get('custom_name', os.path.splitext(file_name)[0])
                        
                        # 确保sheet名称有效
                        sheet_name = self.sanitize_sheet_name(sheet_name)
                        
                        # 保存数据
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 应用样式
                        if merge_config['keep_styles'] and header_styles and data_styles and self.style_manager:
                            wb = writer.book
                            self.style_manager.apply_column_styles(
                                wb,
                                sheet_name,
                                header_styles,
                                data_styles,
                                merge_config
                            )
            
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def read_excel_range(self, file_path, sheet_name, header_row, start_row=None, end_row=None, 
                        start_col=None, end_col=None, add_source=True):
        """读取指定范围的Excel数据"""
        try:
            # 获取表头行号
            header_row = int(header_row) - 1  # 转换为0-based索引
            
            # 处理行列范围
            start_row = int(start_row) - 1 if start_row else None
            end_row = int(end_row) if end_row else None
            start_col = self.col_to_num(start_col) if start_col else None
            end_col = self.col_to_num(end_col) + 1 if end_col else None
            
            # 读取Excel文件，指定表头行
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
            
            # 截取指定范围
            df = df.iloc[start_row:end_row, start_col:end_col]
            
            # 添加数据来源列
            if add_source:
                df['数据来源'] = os.path.basename(file_path)
            
            return df
            
        except Exception as e:
            raise Exception(f"读取文件 {os.path.basename(file_path)} 的 {sheet_name} 时出错: {str(e)}")

    def smart_merge(self, dataframes, keep_header=True):
        """智能合并数据框列表"""
        if not dataframes:
            return pd.DataFrame()
            
        # 获取所有列名（除了'数据来源'）
        all_columns = set()
        for df in dataframes:
            all_columns.update([col for col in df.columns if col != '数据来源'])
            
        # 确保所有数据框都有相同的列
        for df in dataframes:
            for col in all_columns:
                if col not in df.columns:
                    df[col] = None
                    
        # 根据是否保留表头选择合并方式
        if keep_header:
            # 保留表头的合并方式
            result_df = pd.concat(dataframes, ignore_index=True)
        else:
            # 不保留表头的合并方式（跳过第一个文件之后的表头行）
            result_df = dataframes[0].copy()  # 第一个文件完整保留
            
            # 合并其他文件，跳过表头行
            for df in dataframes[1:]:
                result_df = pd.concat([result_df, df], ignore_index=True)
        
        # 调整列顺序，确保'数据来源'列在最后
        if '数据来源' in result_df.columns:
            cols = [col for col in result_df.columns if col != '数据来源'] + ['数据来源']
            result_df = result_df[cols]
        
        return result_df

    def check_headers_consistency(self, dataframes):
        """检查所有数据框的表头是否一致"""
        if not dataframes:
            return False, "没有数据可供检查"
            
        # 获取第一个数据框的列（不包括'数据来源'列）
        base_columns = set(col for col in dataframes[0].columns if col != '数据来源')
        
        # 检查其他数据框的列是否与第一个相同
        inconsistent_files = []
        for i, df in enumerate(dataframes[1:], 1):
            current_columns = set(col for col in df.columns if col != '数据来源')
            if current_columns != base_columns:
                file_name = df['数据来源'].iloc[0]
                diff_cols = base_columns.symmetric_difference(current_columns)
                inconsistent_files.append(f"文件 {file_name} 的列不一致，差异列：{', '.join(diff_cols)}")
                
        if inconsistent_files:
            return False, "\n".join(inconsistent_files)
        return True, "所有文件的表头一致"

    def save_merged_file(self, output_file, data, sheet_name="合并结果", header_styles=None, data_styles=None, merge_config=None):
        """保存合并后的文件"""
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 应用样式
            if merge_config and merge_config['keep_styles'] and header_styles and data_styles and self.style_manager:
                wb = writer.book
                self.style_manager.apply_column_styles(
                    wb,
                    sheet_name,
                    header_styles,
                    data_styles,
                    merge_config
                )

    @staticmethod
    def col_to_num(col_str):
        """将Excel列标转换为数字"""
        if not col_str:
            return None
        num = 0
        for c in col_str.upper():
            num = num * 26 + (ord(c) - ord('A') + 1)
        return num - 1

    @staticmethod
    def sanitize_sheet_name(sheet_name):
        """确保sheet名称有效"""
        # Excel的sheet名称限制：
        # 1. 长度不能超过31个字符
        # 2. 不能包含特殊字符: [ ] : * ? / \
        # 3. 不能为空
        
        # 移除非法字符
        invalid_chars = r'[]*?/\\'
        for char in invalid_chars:
            sheet_name = sheet_name.replace(char, '_')
            
        # 限制长度
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
            
        # 确保不为空
        if not sheet_name:
            sheet_name = "Sheet1"
            
        return sheet_name 