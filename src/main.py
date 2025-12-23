#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX文档提取器主入口
GUI模式
"""

import os
import sys

# 添加src目录到Python路径，支持打包后的exe文件
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 对于PyInstaller打包的exe，需要处理特殊路径
if hasattr(sys, '_MEIPASS'):
    # 这是PyInstaller打包后的临时目录
    bundle_dir = sys._MEIPASS
    sys.path.insert(0, bundle_dir)
else:
    # 开发环境或直接运行Python脚本
    bundle_dir = current_dir

try:
    from gui import main as gui_main
except ImportError as e:
    print(f"导入模块失败: {e}")
    print(f"当前工作目录: {os.getcwd()}")
    print(f"脚本目录: {current_dir}")
    print(f"Python路径: {sys.path}")
    if hasattr(sys, '_MEIPASS'):
        print(f"PyInstaller临时目录: {sys._MEIPASS}")
    sys.exit(1)


def main() -> int:
    """主函数，启动GUI界面
    
    Returns:
        退出码，0表示成功，非0表示失败
    """
    return gui_main()


if __name__ == '__main__':
    sys.exit(main())
