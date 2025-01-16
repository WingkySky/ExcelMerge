"""
Excel样式管理模块
处理Excel文件的样式复制和应用
"""
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter

class ExcelStyleManager:
    def __init__(self, keep_styles=True, keep_colors=True, keep_cell_format=True):
        """初始化Excel样式管理器"""
        self.keep_styles = keep_styles
        self.keep_colors = keep_colors
        self.keep_cell_format = keep_cell_format

    def get_column_styles(self, workbook, sheet_name, header_row):
        """获取列样式模板"""
        sheet = workbook[sheet_name]
        
        # 获取表头行的样式
        header_styles = {}
        for cell in sheet[header_row]:
            if not isinstance(cell, type(None)):
                try:
                    col_letter = get_column_letter(cell.column)
                    if hasattr(cell, 'font'):
                        # 创建新的样式对象而不是直接引用
                        header_styles[col_letter] = {
                            'font': Font(
                                name=cell.font.name,
                                size=cell.font.size,
                                bold=cell.font.bold,
                                italic=cell.font.italic,
                                vertAlign=cell.font.vertAlign,
                                underline=cell.font.underline,
                                strike=cell.font.strike,
                                color=cell.font.color
                            ),
                            'fill': PatternFill(
                                fill_type=cell.fill.fill_type,
                                start_color=cell.fill.start_color,
                                end_color=cell.fill.end_color
                            ),
                            'border': Border(
                                left=cell.border.left,
                                right=cell.border.right,
                                top=cell.border.top,
                                bottom=cell.border.bottom
                            ),
                            'alignment': Alignment(
                                horizontal=cell.alignment.horizontal,
                                vertical=cell.alignment.vertical,
                                text_rotation=cell.alignment.text_rotation,
                                wrap_text=cell.alignment.wrap_text,
                                shrink_to_fit=cell.alignment.shrink_to_fit,
                                indent=cell.alignment.indent
                            ),
                            'number_format': cell.number_format
                        }
                except (AttributeError, TypeError):
                    continue
            
        # 获取数据行的样式（使用表头行下一行作为模板）
        data_styles = {}
        if sheet.max_row > header_row:
            for cell in sheet[header_row + 1]:
                if not isinstance(cell, type(None)):
                    try:
                        col_letter = get_column_letter(cell.column)
                        if hasattr(cell, 'font'):
                            # 创建新的样式对象而不是直接引用
                            data_styles[col_letter] = {
                                'font': Font(
                                    name=cell.font.name,
                                    size=cell.font.size,
                                    bold=cell.font.bold,
                                    italic=cell.font.italic,
                                    vertAlign=cell.font.vertAlign,
                                    underline=cell.font.underline,
                                    strike=cell.font.strike,
                                    color=cell.font.color
                                ),
                                'fill': PatternFill(
                                    fill_type=cell.fill.fill_type,
                                    start_color=cell.fill.start_color,
                                    end_color=cell.fill.end_color
                                ),
                                'border': Border(
                                    left=cell.border.left,
                                    right=cell.border.right,
                                    top=cell.border.top,
                                    bottom=cell.border.bottom
                                ),
                                'alignment': Alignment(
                                    horizontal=cell.alignment.horizontal,
                                    vertical=cell.alignment.vertical,
                                    text_rotation=cell.alignment.text_rotation,
                                    wrap_text=cell.alignment.wrap_text,
                                    shrink_to_fit=cell.alignment.shrink_to_fit,
                                    indent=cell.alignment.indent
                                ),
                                'number_format': cell.number_format
                            }
                    except (AttributeError, TypeError):
                        continue
                
        return header_styles, data_styles

    def apply_column_styles(self, workbook, sheet_name, header_styles, data_styles, merge_config):
        """应用列样式"""
        if not merge_config['keep_styles']:
            return
            
        sheet = workbook[sheet_name]
        
        try:
            # 应用表头样式
            for cell in sheet[1]:  # 第一行是表头
                if not isinstance(cell, type(None)):
                    try:
                        col_letter = get_column_letter(cell.column)
                        if col_letter in header_styles:
                            style = header_styles[col_letter]
                            if merge_config['keep_styles']:
                                cell.font = style['font']
                                cell.border = style['border']
                                cell.alignment = style['alignment']
                            if merge_config['keep_colors']:
                                cell.fill = style['fill']
                            if merge_config['keep_cell_format']:
                                cell.number_format = style['number_format']
                    except (AttributeError, TypeError):
                        continue
                    
            # 应用数据行样式
            for row in sheet.iter_rows(min_row=2):  # 从第二行开始是数据
                for cell in row:
                    if not isinstance(cell, type(None)):
                        try:
                            col_letter = get_column_letter(cell.column)
                            if col_letter in data_styles:
                                style = data_styles[col_letter]
                                if merge_config['keep_styles']:
                                    cell.font = style['font']
                                    cell.border = style['border']
                                    cell.alignment = style['alignment']
                                if merge_config['keep_colors']:
                                    cell.fill = style['fill']
                                if merge_config['keep_cell_format']:
                                    cell.number_format = style['number_format']
                        except (AttributeError, TypeError):
                            continue
                            
            # 调整列宽
            if merge_config['keep_column_width']:
                self.adjust_column_width(sheet)
                
        except Exception as e:
            print(f"应用样式时出错: {str(e)}")
            # 继续执行，即使样式应用失败

    def adjust_column_width(self, worksheet):
        """调整列宽"""
        for column in worksheet.columns:
            max_length = 0
            try:
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        continue
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            except:
                continue 