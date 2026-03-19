import sys
import os

# 添加当前目录到路径，确保能导入其他模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui import MainWindow
from unlock_dialog import UnlockDialog
from extractor import ArchiveExtractor


def main():
    """程序入口"""
    # 检查是否有命令行参数（文件路径）
    if len(sys.argv) > 1:
        file_path = sys.argv[1]

        # 确保文件存在
        if not os.path.exists(file_path):
            print(f"错误: 文件不存在\n{file_path}")
            input("按回车键退出...")
            sys.exit(1)

        # 检查是否是支持的格式
        extractor = ArchiveExtractor()
        if not extractor.is_supported(file_path):
            print(f"错误: 不支持的格式\n{file_path}")
            input("按回车键退出...")
            sys.exit(1)

        # 显示解锁对话框
        dialog = UnlockDialog(file_path)
        dialog.show()

    else:
        # 没有命令行参数，显示主窗口
        app = MainWindow()
        app.run()


if __name__ == "__main__":
    main()