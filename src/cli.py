#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX文档提取器命令行接口
"""

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import List

from config import Config
from core import DocumentExtractor


def setup_logging(verbose: bool = False, quiet: bool = False) -> None:
    """设置日志系统
    
    Args:
        verbose: 是否启用详细输出模式
        quiet: 是否启用静默模式
    """
    if quiet:
        level = logging.ERROR
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO
    
    formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    
    logger = logging.getLogger()
    logger.setLevel(level)
    logger.handlers.clear()
    logger.addHandler(handler)


def find_docx_files(path: str) -> List[str]:
    """查找指定路径下的所有DOCX文件
    
    Args:
        path: 文件或目录路径
        
    Returns:
        DOCX文件路径列表
        
    Raises:
        ValueError: 文件不是DOCX格式
        FileNotFoundError: 路径不存在
    """
    path_obj = Path(path)
    
    if path_obj.is_file():
        if path_obj.suffix.lower() == '.docx':
            return [str(path_obj)]
        else:
            raise ValueError(f"文件 {path} 不是DOCX格式")
    
    elif path_obj.is_dir():
        docx_files: List[str] = []
        for file_path in path_obj.rglob('*.docx'):
            # 跳过临时文件
            if not file_path.name.startswith('~$'):
                docx_files.append(str(file_path))
        return docx_files
    
    else:
        raise FileNotFoundError(f"路径不存在: {path}")


def process_single_file(input_file: str, output_dir: str, 
                        config: Config) -> bool:
    """处理单个文件
    
    Args:
        input_file: 输入文件路径
        output_dir: 输出目录路径
        config: 配置对象
        
    Returns:
        处理是否成功
    """
    try:
        extractor = DocumentExtractor()
        content = extractor.extract_content(input_file)
        
        # 确定输出文件路径
        input_path = Path(input_file)
        if output_dir:
            output_path = Path(output_dir) / f"{input_path.stem}.txt"
        else:
            output_path = input_path.with_suffix('.txt')
        
        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        logging.info(f"处理完成: {input_file} -> {output_path}")
        return True
        
    except Exception as e:
        logging.error(f"处理文件 {input_file} 时出错: {str(e)}")
        return False


def main() -> int:
    """主函数
    
    Returns:
        退出码，0表示成功，非0表示失败
    """
    parser = argparse.ArgumentParser(
        description='DOCX文档提取器 - 将Word文档转换为格式化文本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  %(prog)s document.docx                    # 转换单个文件
  %(prog)s -i docs/ -o output/             # 批量转换目录
  %(prog)s document.docx -o output/        # 指定输出目录
  %(prog)s -i docs/ --verbose              # 详细输出模式
  %(prog)s -i docs/ --quiet                # 静默模式
        """
    )
    
    # 输入输出参数
    parser.add_argument('input', nargs='?', help='输入DOCX文件或目录路径')
    parser.add_argument('-i', '--input', dest='input_path', help='输入DOCX文件或目录路径')
    parser.add_argument('-o', '--output', dest='output_dir', help='输出目录路径（默认为输入文件同目录）')
    
    # 格式控制参数
    parser.add_argument('--width', type=int, default=80, help='文本行宽度（默认80字符）')
    parser.add_argument('--indent', default='    ', help='段落缩进（默认4个空格）')
    
    # 处理模式参数
    parser.add_argument('--batch', action='store_true', help='批量处理模式（递归查找所有DOCX文件）')
    
    # 日志控制参数
    parser.add_argument('-v', '--verbose', action='store_true', help='详细输出模式')
    parser.add_argument('-q', '--quiet', action='store_true', help='静默模式（仅显示错误）')
    
    # 配置参数
    parser.add_argument('--config', help='配置文件路径')
    
    args = parser.parse_args()
    
    # 确定输入路径
    input_path = args.input or args.input_path
    if not input_path:
        parser.error("必须指定输入文件或目录")
    
    # 设置日志
    setup_logging(args.verbose, args.quiet)
    
    # 加载配置
    config = Config()
    if args.config:
        config.load_from_file(args.config)
    
    # 应用命令行参数到配置
    if args.width != 80:
        config.text_width = args.width
    if args.indent != '    ':
        config.text_indent = args.indent
    
    try:
        # 查找要处理的文件
        if args.batch or os.path.isdir(input_path):
            docx_files = find_docx_files(input_path)
            if not docx_files:
                logging.warning(f"在 {input_path} 中未找到DOCX文件")
                return 1
            
            logging.info(f"找到 {len(docx_files)} 个DOCX文件")
        else:
            docx_files = [input_path]
        
        # 处理文件
        success_count = 0
        total_count = len(docx_files)
        
        for file_path in docx_files:
            if process_single_file(file_path, args.output_dir, config):
                success_count += 1
        
        # 输出结果统计
        if total_count > 1:
            logging.info(f"处理完成: {success_count}/{total_count} 个文件成功")
        
        return 0 if success_count == total_count else 1
        
    except KeyboardInterrupt:
        logging.info("用户中断处理")
        return 130
    except Exception as e:
        logging.error(f"程序执行出错: {str(e)}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
