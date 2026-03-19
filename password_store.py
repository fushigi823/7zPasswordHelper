import json
import os
import sys

class PasswordStore:
    """密码存储管理类"""

    def __init__(self, config_path=None):
        if config_path is None:
            if getattr(sys, 'frozen', False):
                app_dir = os.path.dirname(sys.executable)
            else:
                app_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(app_dir, "config.json")
        self.config_path = config_path
        self._ensure_config_exists()

    def _ensure_config_exists(self):
        """确保配置文件存在"""
        if not os.path.exists(self.config_path):
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump({"passwords": {}}, f)

    def _load(self):
        """加载配置"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            return {"passwords": {}}

    def _save(self, data):
        """保存配置"""
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def list(self):
        """列出所有密码别名和密码"""
        data = self._load()
        return data.get("passwords", {})

    def add(self, alias, password):
        """添加或更新密码"""
        data = self._load()
        data["passwords"][alias] = password
        self._save(data)
        return True

    def delete(self, alias):
        """删除密码"""
        data = self._load()
        if alias in data["passwords"]:
            del data["passwords"][alias]
            self._save(data)
            return True
        return False

    def get(self, alias):
        """获取密码"""
        data = self._load()
        return data["passwords"].get(alias)

    def get_aliases(self):
        """获取所有别名列表"""
        return list(self.list().keys())