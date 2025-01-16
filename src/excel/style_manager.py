"""
Excel样式管理模块
处理Excel文件样式的保存和应用
"""
from openpyxl.utils import get_column_letter
from copy import copy

class ExcelStyleManager:
    def __init__(self):
        """初始化样式管理器"""
        pass
        
    def get_column_styles(self, workbook, sheet_name, header_row):
        """
        获取指定sheet的列样式
        
        Args:
            workbook: openpyxl的Workbook对象
            sheet_name: sheet名称
            header_row: 表头行号（1-based）
            
        Returns:
            tuple: (header_styles, data_styles)
        """
        try:
            sheet = workbook[sheet_name]
            header_styles = {}
            data_styles = {}
            
            # 获取表头样式
            header_row_num = int(header_row)
            for col in range(1, sheet.max_column + 1):
                col_letter = get_column_letter(col)
                cell = sheet[f"{col_letter}{header_row_num}"]
                header_styles[col] = {
                    'font': copy(cell.font),
                    'fill': copy(cell.fill),
                    'border': copy(cell.border),
                    'alignment': copy(cell.alignment),
                    'number_format': cell.number_format,
                    'protection': copy(cell.protection)
                }
            
            # 获取数据区域样式（使用第一个数据行的样式作为模板）
            data_row_num = header_row_num + 1
            if data_row_num <= sheet.max_row:
                for col in range(1, sheet.max_column + 1):
                    col_letter = get_column_letter(col)
                    cell = sheet[f"{col_letter}{data_row_num}"]
                    data_styles[col] = {
                        'font': copy(cell.font),
                        'fill': copy(cell.fill),
                        'border': copy(cell.border),
                        'alignment': copy(cell.alignment),
                        'number_format': cell.number_format,
                        'protection': copy(cell.protection)
                    }
            
            return header_styles, data_styles
            
        except Exception as e:
            print(f"获取样式时出错: {str(e)}")
            return None, None
            
    def apply_column_styles(self, workbook, sheet_name, header_styles, data_styles, merge_config):
        """
        应用列样式到指定sheet
        
        Args:
            workbook: openpyxl的Workbook对象
            sheet_name: sheet名称
            header_styles: 表头样式字典
            data_styles: 数据样式字典
            merge_config: 合并配置
        """
        try:
            sheet = workbook[sheet_name]
            
            if merge_config['keep_styles'] and header_styles and data_styles:
                # 应用表头样式
                header_row = int(merge_config['header_row'])
                for col in range(1, sheet.max_column + 1):
                    if col in header_styles:
                        col_letter = get_column_letter(col)
                        cell = sheet[f"{col_letter}{header_row}"]
                        self._apply_cell_style(cell, header_styles[col], merge_config)
                
                # 应用数据区域样式
                for row in range(header_row + 1, sheet.max_row + 1):
                    for col in range(1, sheet.max_column + 1):
                        if col in data_styles:
                            col_letter = get_column_letter(col)
                            cell = sheet[f"{col_letter}{row}"]
                            self._apply_cell_style(cell, data_styles[col], merge_config)
                
                # 调整列宽
                if merge_config['keep_column_width']:
                    self._adjust_column_width(sheet)
                    
        except Exception as e:
            print(f"应用样式时出错: {str(e)}")
            raise  # 抛出异常，让调用者知道出错了
            
    def _apply_cell_style(self, cell, style, merge_config):
        """应用单元格样式"""
        try:
            # 创建新的字体对象，保留所有属性
            new_font = copy(style['font'])
            
            if merge_config['keep_styles']:
                # 如果不保留颜色，清除字体颜色
                if not merge_config['keep_colors']:
                    new_font.color = None
                cell.font = new_font
                cell.border = copy(style['border'])
                cell.alignment = copy(style['alignment'])
                cell.protection = copy(style['protection'])
                
            if merge_config['keep_colors']:
                cell.fill = copy(style['fill'])
                # 如果只保留颜色不保留样式，只复制字体颜色
                if not merge_config['keep_styles'] and style['font'].color:
                    current_font = copy(cell.font)
                    current_font.color = style['font'].color
                    cell.font = current_font
                    
            if merge_config['keep_cell_format']:
                cell.number_format = style['number_format']
                
        except Exception as e:
            print(f"应用单元格样式时出错: {str(e)}")
            # 继续处理，不中断整个过程
            
    def _adjust_column_width(self, sheet):
        """调整列宽"""
        for col in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(col)
            max_length = 0
            for cell in sheet[col_letter]:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[col_letter].width = adjusted_width 