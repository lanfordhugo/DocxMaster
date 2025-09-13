#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX文档提取器主入口
支持CLI和GUI两种模式
"""

import os
import sys

# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cli import main as cli_main
from gui import main as gui_main


def main() -> int:
    """主函数，根据参数决定使用CLI还是GUI
    
    Returns:
        退出码，0表示成功，非0表示失败
    """
    # 如果有命令行参数，使用CLI模式
    if len(sys.argv) > 1:
        return cli_main()
    else:
        # 否则使用GUI模式
        return gui_main()


if __name__ == '__main__':
    sys.exit(main())
