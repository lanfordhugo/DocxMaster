#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置管理模块
"""

from pathlib import Path
from typing import Any, Dict

try:
    import yaml
except ImportError:
    yaml = None


class Config:
    """配置管理类"""
    
    def __init__(self) -> None:
        """初始化默认配置"""
        # 文本格式配置
        self.text_width: int = 80
        self.text_indent: str = "    "
        self.heading_prefix: str = "#"
        
        # 表格格式配置
        self.base_column_width: int = 15
        self.level_2_multiplier: int = 2
        self.level_3_multiplier: int = 3
        self.cell_padding: int = 2
        self.cell_left_padding: int = 1
        
        # 输出配置（固定为md格式）
        self.preserve_structure: bool = True
        self.merge_consecutive_empty_lines: bool = True
        
        # 处理配置
        self.skip_temp_files: bool = True  # 跳过~$开头的临时文件
        self.recursive_search: bool = False
        
    def load_from_file(self, config_path: str) -> None:
        """从配置文件加载配置
        
        Args:
            config_path: 配置文件路径（支持.yaml/.yml/.json格式）
            
        Raises:
            FileNotFoundError: 配置文件不存在
            ValueError: 配置文件格式错误或YAML模块未安装
            RuntimeError: 加载配置文件失败
        """
        config_file = Path(config_path)
        
        if not config_file.exists():
            raise FileNotFoundError(f"配置文件不存在: {config_path}")
        
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                # 根据文件扩展名选择解析方式
                if config_file.suffix.lower() in ['.yaml', '.yml']:
                    if yaml is None:
                        raise ValueError("YAML格式需要安装PyYAML: pip install PyYAML")
                    config_data = yaml.safe_load(f)
                elif config_file.suffix.lower() == '.json':
                    import json
                    config_data = json.load(f)
                else:
                    # 默认尝试YAML格式
                    if yaml is None:
                        raise ValueError("YAML格式需要安装PyYAML: pip install PyYAML")
                    config_data = yaml.safe_load(f)
            
            self._update_from_dict(config_data)
            
        except yaml.YAMLError as e:
            raise ValueError(f"YAML配置文件格式错误: {e}")
        except Exception as e:
            raise RuntimeError(f"加载配置文件失败: {e}")
    
    def save_to_file(self, config_path: str) -> None:
        """保存配置到文件
        
        Args:
            config_path: 配置文件路径（支持.yaml/.yml/.json格式）
            
        Raises:
            ValueError: YAML模块未安装
            RuntimeError: 保存配置文件失败
        """
        config_data = self._to_dict()
        
        config_file = Path(config_path)
        config_file.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                # 根据文件扩展名选择保存格式
                if config_file.suffix.lower() in ['.yaml', '.yml']:
                    if yaml is None:
                        raise ValueError("YAML格式需要安装PyYAML: pip install PyYAML")
                    yaml.dump(config_data, f, default_flow_style=False, 
                             allow_unicode=True, indent=2, sort_keys=False)
                elif config_file.suffix.lower() == '.json':
                    import json
                    json.dump(config_data, f, ensure_ascii=False, indent=2)
                else:
                    # 默认使用YAML格式
                    if yaml is None:
                        raise ValueError("YAML格式需要安装PyYAML: pip install PyYAML")
                    yaml.dump(config_data, f, default_flow_style=False, 
                             allow_unicode=True, indent=2, sort_keys=False)
        except Exception as e:
            raise RuntimeError(f"保存配置文件失败: {e}")
    
    def _update_from_dict(self, config_data: Dict[str, Any]) -> None:
        """从字典更新配置
        
        Args:
            config_data: 配置数据字典
        """
        for key, value in config_data.items():
            if hasattr(self, key):
                setattr(self, key, value)
    
    def _to_dict(self) -> Dict[str, Any]:
        """转换为字典
        
        Returns:
            配置数据字典
        """
        return {
            # 文本格式配置
            'text_width': self.text_width,
            'text_indent': self.text_indent,
            'heading_prefix': self.heading_prefix,
            
            # 表格格式配置
            'base_column_width': self.base_column_width,
            'level_2_multiplier': self.level_2_multiplier,
            'level_3_multiplier': self.level_3_multiplier,
            'cell_padding': self.cell_padding,
            'cell_left_padding': self.cell_left_padding,
            
            # 输出配置（固定为md格式）
            'preserve_structure': self.preserve_structure,
            'merge_consecutive_empty_lines': self.merge_consecutive_empty_lines,
            
            # 处理配置
            'skip_temp_files': self.skip_temp_files,
            'recursive_search': self.recursive_search,
        }
    
    @classmethod
    def create_default_config(cls, config_path: str) -> None:
        """创建默认配置文件
        
        Args:
            config_path: 配置文件路径
        """
        config = cls()
        config.save_to_file(config_path)
