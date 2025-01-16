"""
路径管理模块
处理文件路径的保存和加载
"""
import os

class PathManager:
    def __init__(self, recent_paths_file="recent_paths.txt", max_recent_paths=5):
        """初始化路径管理器"""
        self.recent_paths_file = recent_paths_file
        self.max_recent_paths = max_recent_paths
        self.recent_paths = []
        self.default_paths = [
            os.path.expanduser("~/Documents"),  # 文档文件夹
            os.path.expanduser("~/Desktop"),    # 桌面
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),  # 程序所在目录
        ]
        self.load_recent_paths()

    def load_recent_paths(self):
        """加载最近使用的路径"""
        try:
            if os.path.exists(self.recent_paths_file):
                with open(self.recent_paths_file, 'r', encoding='utf-8') as f:
                    paths = f.read().splitlines()
                self.recent_paths = [p for p in paths if os.path.exists(p)]  # 只返回仍然存在的路径
        except Exception:
            self.recent_paths = []
        return self.recent_paths

    def save_recent_paths(self):
        """保存最近使用的路径"""
        try:
            with open(self.recent_paths_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.recent_paths))
        except Exception:
            pass

    def add_recent_path(self, path):
        """添加新的最近使用路径"""
        if path and os.path.exists(path):
            # 如果路径已存在，先移除它
            if path in self.recent_paths:
                self.recent_paths.remove(path)
            # 添加到列表开头
            self.recent_paths.insert(0, path)
            # 保持列表长度不超过最大值
            if len(self.recent_paths) > self.max_recent_paths:
                self.recent_paths = self.recent_paths[:self.max_recent_paths]
            # 保存更新后的列表
            self.save_recent_paths()

    def get_available_paths(self):
        """获取所有可用路径（包括最近使用的和默认的）"""
        # 合并最近路径和默认路径，去重
        return self.recent_paths + [p for p in self.default_paths if p not in self.recent_paths]

    def get_default_output_path(self):
        """获取默认输出路径"""
        # 优先使用最近使用的路径
        if self.recent_paths:
            return self.recent_paths[0]
        # 否则使用第一个默认路径
        if self.default_paths:
            return self.default_paths[0]
        # 如果都没有，使用当前目录
        return os.getcwd() 