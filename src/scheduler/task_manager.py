"""
任务管理模块
处理定时任务的管理和执行
"""
import os
import json
import threading
import time
from datetime import datetime
from .task_config import TaskConfig

class TaskManager:
    def __init__(self, app):
        """初始化任务管理器"""
        self.app = app
        self.tasks = {}  # {task_id: TaskConfig}
        self.running = False
        self.thread = None
        self.config_dir = os.path.expanduser("~/.excel_merger/tasks")
        self.load_tasks()
        
    def load_tasks(self):
        """加载所有任务配置"""
        try:
            # 确保配置目录存在
            os.makedirs(self.config_dir, exist_ok=True)
            
            # 加载所有任务文件
            for filename in os.listdir(self.config_dir):
                if filename.endswith('.json'):
                    file_path = os.path.join(self.config_dir, filename)
                    task = TaskConfig.load_from_file(file_path)
                    if task:
                        self.tasks[task.task_id] = task
        except Exception as e:
            print(f"加载任务配置失败：{str(e)}")
            
    def save_tasks(self):
        """保存所有任务配置"""
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            for task_id, task in self.tasks.items():
                file_path = os.path.join(self.config_dir, f"{task_id}.json")
                task.save_to_file(file_path)
        except Exception as e:
            print(f"保存任务配置失败：{str(e)}")
            
    def add_task(self, task):
        """添加新任务"""
        self.tasks[task.task_id] = task
        self.save_tasks()
        
    def remove_task(self, task_id):
        """删除任务"""
        if task_id in self.tasks:
            del self.tasks[task_id]
            # 删除配置文件
            file_path = os.path.join(self.config_dir, f"{task_id}.json")
            try:
                os.remove(file_path)
            except:
                pass
            self.save_tasks()
            
    def update_task(self, task):
        """更新任务"""
        if task.task_id in self.tasks:
            self.tasks[task.task_id] = task
            self.save_tasks()
            
    def get_task(self, task_id):
        """获取任务"""
        return self.tasks.get(task_id)
        
    def get_all_tasks(self):
        """获取所有任务"""
        return list(self.tasks.values())
        
    def start(self):
        """启动任务管理器"""
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self._run, daemon=True)
            self.thread.start()
            
    def stop(self):
        """停止任务管理器"""
        self.running = False
        if self.thread:
            self.thread.join()
            
    def _run(self):
        """运行任务检查循环"""
        while self.running:
            now = datetime.now()
            
            # 检查每个任务
            for task in self.tasks.values():
                if task.enabled and task.next_run:
                    next_run = datetime.strptime(task.next_run, "%Y-%m-%d %H:%M:00")
                    if now >= next_run:
                        self._execute_task(task)
                        
            # 每分钟检查一次
            time.sleep(60)
            
    def _execute_task(self, task):
        """执行任务"""
        try:
            # 验证文件
            if not task.validate_files():
                raise Exception("任务包含的文件不存在")
                
            # 执行合并
            self.app.merge_handler.merge_files_with_config(task)
            
            # 更新任务状态
            task.last_run = datetime.now().strftime("%Y-%m-%d %H:%M:00")
            task.update_next_run()
            self.save_tasks()
            
        except Exception as e:
            print(f"执行任务失败：{str(e)}")
            # 可以添加错误通知机制 