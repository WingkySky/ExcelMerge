"""
路径管理模块
处理输出路径的管理和记录
"""
import os
import json

class PathManager:
    def __init__(self):
        """初始化路径管理器"""
        self.config_file = os.path.expanduser("~/.excel_merger/paths.json")
        self.recent_paths = []
        self.max_paths = 5
        self.load_paths()
        
    def load_paths(self):
        """加载保存的路径"""
        try:
            # 确保配置目录存在
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            # 如果配置文件存在，读取它
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.recent_paths = json.load(f)
            else:
                # 添加默认路径
                default_paths = [
                    os.path.expanduser("~/Documents"),
                    os.path.expanduser("~/Desktop"),
                    os.path.expanduser("~/Downloads")
                ]
                self.recent_paths = [p for p in default_paths if os.path.exists(p)]
                self.save_paths()
        except Exception:
            # 如果出现任何错误，使用默认路径
            self.recent_paths = [os.path.expanduser("~/Documents")]
            
    def save_paths(self):
        """保存路径到配置文件"""
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.recent_paths, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
            
    def add_recent_path(self, path):
        """添加新的路径到最近使用列表"""
        if path in self.recent_paths:
            # 如果路径已存在，将其移到最前面
            self.recent_paths.remove(path)
        self.recent_paths.insert(0, path)
        
        # 保持列表在最大长度以内
        self.recent_paths = self.recent_paths[:self.max_paths]
        
        # 保存更新后的路径
        self.save_paths()
        
    def get_available_paths(self):
        """获取可用的路径列表"""
        # 过滤掉不存在的路径
        available_paths = [p for p in self.recent_paths if os.path.exists(p)]
        if not available_paths:
            # 如果没有可用路径，返回文档目录
            return [os.path.expanduser("~/Documents")]
        return available_paths 