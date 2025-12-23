#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX文档提取器打包脚本
使用PyInstaller将Python项目打包为独立可执行文件
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def print_banner(message: str) -> None:
    """打印横幅信息
    
    Args:
        message: 要显示的信息
    """
    print("=" * 50)
    print(message)
    print("=" * 50)


def check_python() -> bool:
    """检查Python是否可用
    
    Returns:
        True如果Python可用，否则False
    """
    try:
        result = subprocess.run([sys.executable, "--version"], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print(f"Python版本: {result.stdout.strip()}")
            return True
        else:
            print("错误: Python不可用")
            return False
    except Exception as e:
        print(f"错误: 检查Python版本失败: {e}")
        return False


def check_project_structure() -> bool:
    """检查项目结构是否正确
    
    Returns:
        True如果项目结构正确，否则False
    """
    main_py = Path("src/main.py")
    if not main_py.exists():
        print("错误: src/main.py文件不存在")
        print("请在项目根目录运行此脚本")
        return False
    
    requirements_txt = Path("requirements.txt")
    if not requirements_txt.exists():
        print("警告: requirements.txt文件不存在")
    
    return True


def clean_build_artifacts() -> None:
    """清理之前的构建产物"""
    print("清理之前的构建产物...")
    
    # 清理目录
    for dir_name in ["build", "__pycache__"]:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name, ignore_errors=True)
            print(f"已删除: {dir_name}")
    
    # 清理文件
    for file_pattern in ["*.spec"]:
        for file_path in Path(".").glob(file_pattern):
            file_path.unlink(missing_ok=True)
            print(f"已删除: {file_path}")
    
    # 清理dist中的exe文件
    exe_file = Path("dist/docx_extractor.exe")
    if exe_file.exists():
        exe_file.unlink()
        print(f"已删除: {exe_file}")
    
    # 递归清理Python缓存
    for cache_dir in Path(".").rglob("__pycache__"):
        shutil.rmtree(cache_dir, ignore_errors=True)
    
    for pyc_file in Path(".").rglob("*.pyc"):
        pyc_file.unlink(missing_ok=True)


def install_requirements() -> bool:
    """安装项目依赖
    
    Returns:
        True如果安装成功，否则False
    """
    requirements_file = Path("requirements.txt")
    if not requirements_file.exists():
        print("跳过依赖安装: requirements.txt不存在")
        return True
    
    print("安装项目依赖...")
    try:
        result = subprocess.run([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ], check=True)
        print("依赖安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"错误: 依赖安装失败: {e}")
        return False


def build_executable() -> bool:
    """构建可执行文件
    
    Returns:
        True如果构建成功，否则False
    """
    print("构建可执行文件...")
    
    # PyInstaller命令参数
    icon_path = Path("src/assets/app_icon.ico")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 打包为单个文件
        "--name", "docx_extractor",     # 可执行文件名
        "--clean",                      # 清理缓存
        "--windowed",                   # Windows下隐藏控制台窗口（无黑色终端样式）
        "--add-data", "src/gui.py;.",   # 添加GUI模块  
        "--add-data", "src/config.py;.", # 添加配置模块
        "--add-data", "src/core.py;.",  # 添加核心模块
        "--add-data", "src/docx_extractor.py;.", # 添加提取器模块
        "--add-data", "src/assets;assets",  # 添加资源目录（包含图标）
        "--paths", "src",               # 添加模块搜索路径
        "src/main.py"                   # 主入口文件
    ]
    
    # 如果图标文件存在，添加到exe图标
    if icon_path.exists():
        cmd.insert(cmd.index("--windowed") + 1, "--icon")
        cmd.insert(cmd.index("--icon") + 1, str(icon_path))
    
    try:
        result = subprocess.run(cmd, check=True)
        print("构建成功!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"错误: PyInstaller构建失败: {e}")
        return False


def check_executable() -> bool:
    """检查生成的可执行文件
    
    Returns:
        True如果可执行文件存在，否则False
    """
    exe_file = Path("dist/docx_extractor.exe")
    if exe_file.exists():
        file_size = exe_file.stat().st_size
        print(f"可执行文件已生成: {exe_file}")
        print(f"文件大小: {file_size:,} 字节")
        print()
        print("使用方法: docx_extractor.exe [docx_file_path]")
        return True
    else:
        print("错误: 可执行文件未在dist目录中找到")
        return False


def final_cleanup() -> None:
    """最终清理，只保留exe文件"""
    print("清理构建产物...")
    
    # 清理build目录
    if Path("build").exists():
        shutil.rmtree("build", ignore_errors=True)
    
    # 清理spec文件
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink(missing_ok=True)
    
    # 清理Python缓存
    for cache_dir in Path(".").rglob("__pycache__"):
        shutil.rmtree(cache_dir, ignore_errors=True)
    
    for pyc_file in Path(".").rglob("*.pyc"):
        pyc_file.unlink(missing_ok=True)


def main() -> int:
    """主函数
    
    Returns:
        退出码，0表示成功，非0表示失败
    """
    print_banner("构建DOCX文档提取器")
    
    # 1. 检查Python环境
    if not check_python():
        return 1
    
    # 2. 检查项目结构
    if not check_project_structure():
        return 1
    
    # 3. 清理之前的构建产物
    clean_build_artifacts()
    
    # 4. 安装依赖
    if not install_requirements():
        return 1
    
    # 5. 构建可执行文件
    if not build_executable():
        return 1
    
    # 6. 检查构建结果
    if not check_executable():
        return 1
    
    # 7. 最终清理
    final_cleanup()
    
    print_banner("构建完成!")
    print("只保留 dist/docx_extractor.exe 文件")
    
    return 0


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n构建被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n构建过程中发生未预期的错误: {e}")
        sys.exit(1)
