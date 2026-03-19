import os
import subprocess
import py7zr
import zipfile
from pathlib import Path

class ArchiveExtractor:
    """压缩包解压类"""

    SUPPORTED_FORMATS = ['.7z', '.zip', '.rar', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2']

    def __init__(self):
        self._7z_path = self._find_7z()

    def _find_7z(self):
        """查找7z命令行工具路径"""
        # 检查常见路径
        common_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path

        # 尝试从PATH中查找
        import shutil
        found = shutil.which('7z')
        if found:
            return found

        return '7z'  # 默认假设在PATH中

    def is_supported(self, file_path):
        """检查文件是否是支持的压缩格式"""
        ext = os.path.splitext(file_path)[1].lower()
        return ext in self.SUPPORTED_FORMATS

    def is_encrypted(self, file_path):
        """检测压缩包是否加密"""
        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.7z':
                return self._is_7z_encrypted(file_path)
            elif ext == '.zip':
                return self._is_zip_encrypted(file_path)
            else:
                # 其他格式默认尝试检测
                return self._is_7z_encrypted(file_path)
        except Exception as e:
            print(f"检测加密状态时出错: {e}")
            return True  # 有密码的可能性，假设加密

    def _is_7z_encrypted(self, file_path):
        """检测7z文件是否加密"""
        try:
            with py7zr.SevenZipFile(file_path, 'r') as z:
                # 尝试读取文件列表，如果加密会失败
                z.read()
            return False
        except Exception as e:
            error_str = str(e).lower()
            if 'password' in error_str or 'encrypted' in error_str or 'secret' in error_str:
                return True
            return False

    def _is_zip_encrypted(self, file_path):
        """检测zip文件是否加密"""
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                for info in z.infolist():
                    if info.flag_bits & 0x1:  # 如果设置了加密标志
                        return True
            return False
        except Exception as e:
            error_str = str(e).lower()
            if 'password' in error_str or 'encrypted' in error_str:
                return True
            return False

    def get_default_output_dir(self, file_path):
        """获取默认解压目录（与压缩包同目录下的同名文件夹）"""
        base_dir = os.path.dirname(os.path.abspath(file_path))
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(base_dir, file_name)
        return output_dir

    def extract_with_password(self, file_path, password, output_dir=None):
        """使用指定密码解压压缩包"""
        if output_dir is None:
            output_dir = self.get_default_output_dir(file_path)

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        ext = os.path.splitext(file_path)[1].lower()

        # 使用7z命令行解压
        cmd = [
            self._7z_path,
            'x',
            file_path,
            f'-o{output_dir}',
            f'-p{password}',
            '-y'  # 全部选是
        ]

        print(f"DEBUG extract: cmd={cmd}")

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300  # 5分钟超时
            )

            if result.returncode == 0:
                return True, output_dir, None
            else:
                error_msg = result.stderr or "解压失败"
                return False, None, error_msg

        except subprocess.TimeoutExpired:
            return False, None, "解压超时"
        except Exception as e:
            return False, None, str(e)

    def try_passwords(self, file_path, passwords):
        """尝试多个密码，找到能成功解压的密码"""
        for alias, password in passwords.items():
            success, output_dir, error = self.extract_with_password(file_path, password)
            if success:
                return True, password, output_dir, None
        return False, None, None, "所有密码都无法解压此文件"