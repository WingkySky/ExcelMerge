"""
定时任务管理模块
处理Excel合并的定时执行
"""
import schedule
import time
import threading

class TaskScheduler:
    def __init__(self):
        """初始化定时任务管理器"""
        self.is_scheduling = False
        self.schedule_time = "00:00"
        self.schedule_thread = None
        self.task = None

    def set_task(self, task):
        """设置要执行的任务"""
        self.task = task

    def set_schedule_time(self, time_str):
        """设置定时执行时间"""
        try:
            # 验证时间格式
            time_parts = time_str.split(":")
            if len(time_parts) != 2:
                raise ValueError("时间格式错误")
                
            hour = int(time_parts[0])
            minute = int(time_parts[1])
            
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError("时间范围错误")
                
            self.schedule_time = f"{hour:02d}:{minute:02d}"
            return True
        except ValueError:
            return False

    def start_schedule(self):
        """启动定时任务"""
        if not self.task:
            raise ValueError("未设置任务")
            
        schedule.clear()
        schedule.every().day.at(self.schedule_time).do(self.task)
        self.is_scheduling = True
        self.schedule_thread = threading.Thread(target=self.run_schedule, daemon=True)
        self.schedule_thread.start()

    def stop_schedule(self):
        """停止定时任务"""
        self.is_scheduling = False
        schedule.clear()
        if self.schedule_thread:
            self.schedule_thread.join(timeout=1)
            self.schedule_thread = None

    def run_schedule(self):
        """运行定时任务循环"""
        while self.is_scheduling:
            schedule.run_pending()
            time.sleep(30)  # 每30秒检查一次

    def is_running(self):
        """检查定时任务是否正在运行"""
        return self.is_scheduling

    def get_schedule_time(self):
        """获取当前设置的定时时间"""
        return self.schedule_time 