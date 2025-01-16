"""
定时任务配置模块
处理定时任务的配置信息
"""
import json
import os
from datetime import datetime

class TaskConfig:
    def __init__(self, task_id=None):
        """初始化任务配置"""
        self.task_id = task_id or datetime.now().strftime("%Y%m%d_%H%M%S")
        self.task_name = ""
        self.schedule_time = "00:00"  # 24小时制
        self.enabled = False
        self.last_run = None
        self.next_run = None
        
        # 文件相关配置
        self.input_files = []  # [(文件路径, 选中的sheet)]
        self.output_path = ""
        self.output_filename = ""
        
        # 合并相关配置
        self.merge_config = {
            'merge_mode': "single",
            'sheet_name_mode': "auto",
            'custom_sheet_name': "Sheet1",
            'start_row': "1",
            'end_row': "",
            'start_col': "A",
            'end_col': "",
            'header_row': "1",
            'keep_header': True,
            'keep_styles': True,
            'keep_column_width': True,
            'keep_cell_format': True,
            'keep_colors': True
        }
        
    def to_dict(self):
        """将配置转换为字典"""
        return {
            'task_id': self.task_id,
            'task_name': self.task_name,
            'schedule_time': self.schedule_time,
            'enabled': self.enabled,
            'last_run': self.last_run,
            'next_run': self.next_run,
            'input_files': self.input_files,
            'output_path': self.output_path,
            'output_filename': self.output_filename,
            'merge_config': self.merge_config
        }
        
    @classmethod
    def from_dict(cls, data):
        """从字典创建配置"""
        task = cls(data.get('task_id'))
        task.task_name = data.get('task_name', '')
        task.schedule_time = data.get('schedule_time', '00:00')
        task.enabled = data.get('enabled', False)
        task.last_run = data.get('last_run')
        task.next_run = data.get('next_run')
        task.input_files = data.get('input_files', [])
        task.output_path = data.get('output_path', '')
        task.output_filename = data.get('output_filename', '')
        task.merge_config = data.get('merge_config', {})
        return task
        
    def save_to_file(self, file_path):
        """保存配置到文件"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存任务配置失败：{str(e)}")
            return False
            
    @classmethod
    def load_from_file(cls, file_path):
        """从文件加载配置"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except Exception as e:
            print(f"加载任务配置失败：{str(e)}")
            return None
            
    def validate_files(self):
        """验证文件是否有效"""
        valid_files = []
        for file_path, sheet in self.input_files:
            if os.path.exists(file_path):
                valid_files.append((file_path, sheet))
        self.input_files = valid_files
        return len(valid_files) > 0
        
    def update_next_run(self):
        """更新下次运行时间"""
        now = datetime.now()
        time_parts = self.schedule_time.split(':')
        next_run = now.replace(
            hour=int(time_parts[0]),
            minute=int(time_parts[1]),
            second=0,
            microsecond=0
        )
        
        # 如果当前时间已经过了今天的执行时间，设置为明天
        if next_run <= now:
            next_run = next_run.replace(day=next_run.day + 1)
            
        self.next_run = next_run.strftime("%Y-%m-%d %H:%M:00") 