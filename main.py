import sys
import os
import logging
import argparse
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
import qt_material

# 导入自定义模块
from ui import MainWindow
from utils import setup_logging


def main():
    """主程序入口"""
    # 设置参数解析
    parser = argparse.ArgumentParser(description="漫画文件夹批量压缩工具")
    parser.add_argument('--dir', type=str, help='要处理的漫画根目录')
    parser.add_argument('--log', type=str, default='comic_compressor.log', help='日志文件路径')
    parser.add_argument('--debug', action='store_true', help='启用调试日志')
    parser.add_argument('--theme', type=str, default='light_blue.xml', help='UI主题')
    args = parser.parse_args()

    # 设置日志
    log_level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(args.log, console_level=log_level)

    logger = logging.getLogger("ComicCompressor")
    logger.info("漫画文件夹批量压缩工具启动")

    # 创建应用程序实例
    app = QApplication(sys.argv)
    app.setApplicationName("漫画文件夹批量压缩工具")

    # 设置应用程序图标
    icon_path = os.path.join(os.path.dirname(__file__), 'resources', 'icon.png')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # 应用材料设计主题
    available_themes = [
        'light_blue.xml',
        'light_cyan.xml',
        'light_lightgreen.xml',
        'light_pink.xml',
        'light_purple.xml',
        'light_red.xml',
        'light_teal.xml',
        'light_yellow.xml',
        'dark_blue.xml',
        'dark_cyan.xml',
        'dark_lightgreen.xml',
        'dark_pink.xml',
        'dark_purple.xml',
        'dark_red.xml',
        'dark_teal.xml',
        'dark_yellow.xml'
    ]

    theme = args.theme if args.theme in available_themes else 'light_blue.xml'
    qt_material.apply_stylesheet(app, theme=theme)

    # 创建并显示主窗口
    window = MainWindow()
    window.show()

    # 如果命令行指定了目录，设置根目录
    if args.dir:
        if os.path.isdir(args.dir):
            window.root_path = args.dir
            window.dir_path_edit.setText(args.dir)
            window.refresh_preview()
        else:
            logger.warning(f"指定的目录不存在: {args.dir}")

    # 执行应用程序
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()