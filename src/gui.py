#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX文档提取器图形界面
"""

import logging
import os
import threading
from pathlib import Path
from typing import List, Optional
import sys

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from config import Config
from core import DocumentExtractor
from src import __version__, __app_name__


class GUILogHandler(logging.Handler):
    """GUI日志处理器"""
    
    def __init__(self, text_widget: tk.Text) -> None:
        """初始化GUI日志处理器
        
        Args:
            text_widget: 用于显示日志的文本框组件
        """
        super().__init__()
        self.text_widget: tk.Text = text_widget
        self.buffer: List[str] = []
        self.update_pending: bool = False
        
    def emit(self, record: logging.LogRecord) -> None:
        """输出日志记录到文本框
        
        Args:
            record: 日志记录对象
        """
        try:
            msg = self.format(record)
            self.buffer.append(msg + '\n')
            
            if not self.update_pending:
                self.update_pending = True
                self.text_widget.after(100, self._flush_buffer)
        except Exception:
            pass
            
    def _flush_buffer(self) -> None:
        """批量更新文本框"""
        try:
            if self.buffer:
                text = ''.join(self.buffer)
                self.text_widget.insert(tk.END, text)
                self.text_widget.see(tk.END)
                self.buffer.clear()
            self.update_pending = False
        except Exception:
            self.update_pending = False


class DocumentExtractorGUI:
    """DOCX提取器图形界面"""
    
    def __init__(self) -> None:
        """初始化GUI界面"""
        self.root: tk.Tk = tk.Tk()
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("1000x750")
        self.root.resizable(True, True)
        
        # 设置窗口图标
        self._set_window_icon()
        
        # 现代化配色方案
        self.colors = {
            'bg': '#f5f5f5',           # 浅灰背景
            'fg': '#2c3e50',           # 深蓝灰文字
            'primary': '#3498db',      # 主色调蓝
            'success': '#27ae60',      # 成功绿
            'danger': '#e74c3c',       # 危险红
            'secondary': '#95a5a6',    # 次要灰
            'card_bg': '#ffffff',      # 卡片白色背景
            'border': '#dcdde1',       # 边框颜色
            'hover': '#2980b9'         # 悬停蓝
        }
        
        # 设置窗口背景色
        self.root.configure(bg=self.colors['bg'])
        
        self.config: Config = Config()
        self.processing: bool = False
        
        # GUI组件
        self.file_path_var: Optional[tk.StringVar] = None
        self.file_entry: Optional[tk.Entry] = None
        self.output_dir_var: Optional[tk.StringVar] = None
        self.output_entry: Optional[tk.Entry] = None
        self.process_btn: Optional[tk.Button] = None
        self.open_output_btn: Optional[tk.Button] = None
        self.log_text: Optional[scrolledtext.ScrolledText] = None
        self.status_var: Optional[tk.StringVar] = None
        
        # 批量处理相关组件
        self.batch_input_var: Optional[tk.StringVar] = None
        self.batch_input_entry: Optional[tk.Entry] = None
        self.batch_output_var: Optional[tk.StringVar] = None
        self.batch_output_entry: Optional[tk.Entry] = None
        self.recursive_var: Optional[tk.BooleanVar] = None
        self.progress_var: Optional[tk.StringVar] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.batch_process_btn: Optional[tk.Button] = None
        self.open_batch_output_btn: Optional[tk.Button] = None
        self.batch_log_text: Optional[scrolledtext.ScrolledText] = None
        
        # 设置相关组件
        self.width_var: Optional[tk.IntVar] = None
        self.indent_var: Optional[tk.StringVar] = None
        self.col_width_var: Optional[tk.IntVar] = None
        
        # 处理结果
        self.last_output_file: Optional[str] = None
        self.batch_output_dir: Optional[str] = None
        
        self.setup_ui()
        self.setup_logging()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            # 确定图标路径（支持打包后和开发环境）
            if getattr(sys, 'frozen', False):
                # PyInstaller 打包后
                base_path = sys._MEIPASS
            else:
                # 开发环境
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, 'assets', 'app_icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass  # 图标加载失败不影响程序运行
        
    def setup_ui(self) -> None:
        """设置用户界面"""
        # 主框架
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 创建笔记本控件（标签页）
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 文件处理标签页
        self.setup_file_tab(notebook)
        
        # 批量处理标签页
        self.setup_batch_tab(notebook)
        
        # 设置标签页
        self.setup_settings_tab(notebook)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = tk.Label(
            self.root, 
            textvariable=self.status_var, 
            relief=tk.FLAT,
            anchor=tk.W,
            bg=self.colors['primary'],
            fg='white',
            font=('Microsoft YaHei UI', 9),
            padx=10,
            pady=5
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def setup_file_tab(self, notebook):
        """设置文件处理标签页"""
        file_frame = ttk.Frame(notebook)
        notebook.add(file_frame, text="📄 单文件处理")
        
        # 文件选择区域 - 卡片样式
        input_card = tk.Frame(file_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        input_card.pack(fill=tk.X, pady=(5, 10), padx=5)
        
        input_frame = tk.Frame(input_card, bg=self.colors['card_bg'])
        input_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(input_frame, text="选择DOCX文件:", bg=self.colors['card_bg'], 
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        self.file_path_var = tk.StringVar()
        self.file_entry = tk.Entry(input_frame, textvariable=self.file_path_var, 
                                   font=('Microsoft YaHei UI', 9), relief=tk.SOLID, bd=1)
        self.file_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True, ipady=5)
        
        browse_btn = tk.Button(input_frame, text="浏览", command=self.browse_file,
                              bg=self.colors['primary'], fg='white', 
                              font=('Microsoft YaHei UI', 9, 'bold'),
                              relief=tk.FLAT, cursor='hand2', padx=20, pady=5)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_hover_effect(browse_btn, self.colors['primary'], self.colors['hover'])
        
        # 输出目录选择 - 卡片样式
        output_card = tk.Frame(file_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        output_card.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        output_frame = tk.Frame(output_card, bg=self.colors['card_bg'])
        output_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(output_frame, text="输出目录:", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar()
        self.output_entry = tk.Entry(output_frame, textvariable=self.output_dir_var,
                                    font=('Microsoft YaHei UI', 9), relief=tk.SOLID, bd=1,
                                    fg='#999999')
        self.output_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True, ipady=5)
        # 添加占位符提示
        self._add_placeholder(self.output_entry, self.output_dir_var, '（可选）不选择则输出到输入文件同级目录')
        
        output_browse_btn = tk.Button(output_frame, text="浏览", command=self.browse_output_dir,
                                     bg=self.colors['primary'], fg='white',
                                     font=('Microsoft YaHei UI', 9, 'bold'),
                                     relief=tk.FLAT, cursor='hand2', padx=20, pady=5)
        output_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_hover_effect(output_browse_btn, self.colors['primary'], self.colors['hover'])
        
        # 处理按钮
        button_frame = tk.Frame(file_frame, bg=self.colors['bg'])
        button_frame.pack(fill=tk.X, pady=(0, 15), padx=5)
        
        self.process_btn = tk.Button(button_frame, text="▶ 开始处理", command=self.process_file,
                                    bg=self.colors['success'], fg='white',
                                    font=('Microsoft YaHei UI', 10, 'bold'),
                                    relief=tk.FLAT, cursor='hand2', padx=30, pady=10)
        self.process_btn.pack(side=tk.LEFT)
        self._add_hover_effect(self.process_btn, self.colors['success'], '#229954')
        
        self.open_output_btn = tk.Button(button_frame, text="📂 打开输出文件", 
                                        command=self.open_output_file, state=tk.DISABLED,
                                        bg=self.colors['secondary'], fg='white',
                                        font=('Microsoft YaHei UI', 9, 'bold'),
                                        relief=tk.FLAT, cursor='hand2', padx=20, pady=8)
        self.open_output_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 日志显示区域 - 卡片样式
        log_card = tk.Frame(file_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        log_card.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        log_header = tk.Frame(log_card, bg=self.colors['card_bg'])
        log_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tk.Label(log_header, text="📋 处理日志", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(anchor=tk.W)
        
        log_content = tk.Frame(log_card, bg=self.colors['card_bg'])
        log_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        self.log_text = scrolledtext.ScrolledText(
            log_content, 
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg='#fafafa',
            fg=self.colors['fg'],
            relief=tk.SOLID,
            bd=1
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def setup_batch_tab(self, notebook):
        """设置批量处理标签页"""
        batch_frame = ttk.Frame(notebook)
        notebook.add(batch_frame, text="📁 批量处理")
        
        # 目录选择区域 - 卡片样式
        input_card = tk.Frame(batch_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        input_card.pack(fill=tk.X, pady=(5, 10), padx=5)
        
        input_frame = tk.Frame(input_card, bg=self.colors['card_bg'])
        input_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(input_frame, text="选择输入目录:", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        self.batch_input_var = tk.StringVar()
        self.batch_input_entry = tk.Entry(input_frame, textvariable=self.batch_input_var,
                                         font=('Microsoft YaHei UI', 9), relief=tk.SOLID, bd=1)
        self.batch_input_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True, ipady=5)
        
        batch_browse_btn = tk.Button(input_frame, text="浏览", command=self.browse_batch_input,
                                    bg=self.colors['primary'], fg='white',
                                    font=('Microsoft YaHei UI', 9, 'bold'),
                                    relief=tk.FLAT, cursor='hand2', padx=20, pady=5)
        batch_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_hover_effect(batch_browse_btn, self.colors['primary'], self.colors['hover'])
        
        # 输出目录选择 - 卡片样式
        output_card = tk.Frame(batch_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        output_card.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        output_frame = tk.Frame(output_card, bg=self.colors['card_bg'])
        output_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(output_frame, text="输出目录:", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        self.batch_output_var = tk.StringVar()
        self.batch_output_entry = tk.Entry(output_frame, textvariable=self.batch_output_var,
                                          font=('Microsoft YaHei UI', 9), relief=tk.SOLID, bd=1,
                                          fg='#999999')
        self.batch_output_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True, ipady=5)
        # 添加占位符提示
        self._add_placeholder(self.batch_output_entry, self.batch_output_var, '（可选）不选择则输出到输入文件同级目录')
        
        batch_output_browse_btn = tk.Button(output_frame, text="浏览", command=self.browse_batch_output,
                                           bg=self.colors['primary'], fg='white',
                                           font=('Microsoft YaHei UI', 9, 'bold'),
                                           relief=tk.FLAT, cursor='hand2', padx=20, pady=5)
        batch_output_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_hover_effect(batch_output_browse_btn, self.colors['primary'], self.colors['hover'])
        
        # 选项区域 - 卡片样式
        options_card = tk.Frame(batch_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        options_card.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        options_frame = tk.Frame(options_card, bg=self.colors['card_bg'])
        options_frame.pack(fill=tk.X, padx=15, pady=12)
        
        self.recursive_var = tk.BooleanVar(value=True)
        recursive_check = tk.Checkbutton(options_frame, text="递归搜索子目录", variable=self.recursive_var,
                                        bg=self.colors['card_bg'], fg=self.colors['fg'],
                                        font=('Microsoft YaHei UI', 9), selectcolor=self.colors['card_bg'])
        recursive_check.pack(side=tk.LEFT)
        
        # 进度条 - 卡片样式
        progress_card = tk.Frame(batch_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        progress_card.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        progress_frame = tk.Frame(progress_card, bg=self.colors['card_bg'])
        progress_frame.pack(fill=tk.X, padx=15, pady=12)
        
        tk.Label(progress_frame, text="处理进度:", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 9, 'bold')).pack(side=tk.LEFT)
        self.progress_var = tk.StringVar(value="0/0")
        tk.Label(progress_frame, textvariable=self.progress_var, bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 9)).pack(side=tk.LEFT, padx=(10, 0))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        # 批量处理按钮
        batch_button_frame = tk.Frame(batch_frame, bg=self.colors['bg'])
        batch_button_frame.pack(fill=tk.X, pady=(0, 15), padx=5)
        
        self.batch_process_btn = tk.Button(batch_button_frame, text="▶ 开始批量处理", command=self.process_batch,
                                          bg=self.colors['success'], fg='white',
                                          font=('Microsoft YaHei UI', 10, 'bold'),
                                          relief=tk.FLAT, cursor='hand2', padx=30, pady=10)
        self.batch_process_btn.pack(side=tk.LEFT)
        self._add_hover_effect(self.batch_process_btn, self.colors['success'], '#229954')
        
        self.open_batch_output_btn = tk.Button(batch_button_frame, text="📂 打开输出目录", 
                                              command=self.open_batch_output, state=tk.DISABLED,
                                              bg=self.colors['secondary'], fg='white',
                                              font=('Microsoft YaHei UI', 9, 'bold'),
                                              relief=tk.FLAT, cursor='hand2', padx=20, pady=8)
        self.open_batch_output_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 批量处理日志 - 卡片样式
        batch_log_card = tk.Frame(batch_frame, bg=self.colors['card_bg'], relief=tk.FLAT, bd=1)
        batch_log_card.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        batch_log_header = tk.Frame(batch_log_card, bg=self.colors['card_bg'])
        batch_log_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tk.Label(batch_log_header, text="📋 批量处理日志", bg=self.colors['card_bg'],
                fg=self.colors['fg'], font=('Microsoft YaHei UI', 10, 'bold')).pack(anchor=tk.W)
        
        batch_log_content = tk.Frame(batch_log_card, bg=self.colors['card_bg'])
        batch_log_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        self.batch_log_text = scrolledtext.ScrolledText(
            batch_log_content, 
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg='#fafafa',
            fg=self.colors['fg'],
            relief=tk.SOLID,
            bd=1
        )
        self.batch_log_text.pack(fill=tk.BOTH, expand=True)
        
    def setup_settings_tab(self, notebook):
        """设置配置标签页"""
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="⚙️ 设置")
        
        # 文本格式设置
        text_group = ttk.LabelFrame(settings_frame, text="文本格式设置")
        text_group.pack(fill=tk.X, padx=5, pady=5)
        
        # 文本宽度
        width_frame = tk.Frame(text_group)
        width_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(width_frame, text="文本行宽度:").pack(side=tk.LEFT)
        self.width_var = tk.IntVar(value=self.config.text_width)
        width_spin = tk.Spinbox(width_frame, from_=40, to=200, textvariable=self.width_var, width=10)
        width_spin.pack(side=tk.LEFT, padx=(10, 0))
        
        # 缩进设置
        indent_frame = tk.Frame(text_group)
        indent_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(indent_frame, text="段落缩进:").pack(side=tk.LEFT)
        self.indent_var = tk.StringVar(value=self.config.text_indent)
        indent_entry = tk.Entry(indent_frame, textvariable=self.indent_var, width=10)
        indent_entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 表格格式设置
        table_group = ttk.LabelFrame(settings_frame, text="表格格式设置")
        table_group.pack(fill=tk.X, padx=5, pady=5)
        
        # 基础列宽
        col_width_frame = tk.Frame(table_group)
        col_width_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(col_width_frame, text="基础列宽:").pack(side=tk.LEFT)
        self.col_width_var = tk.IntVar(value=self.config.base_column_width)
        col_width_spin = tk.Spinbox(col_width_frame, from_=8, to=50, textvariable=self.col_width_var, width=10)
        col_width_spin.pack(side=tk.LEFT, padx=(10, 0))
        
        # 按钮区域
        button_group = tk.Frame(settings_frame, bg=self.colors['bg'])
        button_group.pack(fill=tk.X, padx=5, pady=10)
        
        save_config_btn = tk.Button(button_group, text="💾 保存配置", command=self.save_config,
                                   bg=self.colors['success'], fg='white',
                                   font=('Microsoft YaHei UI', 9, 'bold'),
                                   relief=tk.FLAT, cursor='hand2', padx=20, pady=8)
        save_config_btn.pack(side=tk.LEFT)
        self._add_hover_effect(save_config_btn, self.colors['success'], '#229954')
        
        load_config_btn = tk.Button(button_group, text="📂 加载配置", command=self.load_config,
                                   bg=self.colors['primary'], fg='white',
                                   font=('Microsoft YaHei UI', 9, 'bold'),
                                   relief=tk.FLAT, cursor='hand2', padx=20, pady=8)
        load_config_btn.pack(side=tk.LEFT, padx=(10, 0))
        self._add_hover_effect(load_config_btn, self.colors['primary'], self.colors['hover'])
        
        reset_config_btn = tk.Button(button_group, text="🔄 重置为默认", command=self.reset_config,
                                    bg=self.colors['secondary'], fg='white',
                                    font=('Microsoft YaHei UI', 9, 'bold'),
                                    relief=tk.FLAT, cursor='hand2', padx=20, pady=8)
        reset_config_btn.pack(side=tk.LEFT, padx=(10, 0))
        self._add_hover_effect(reset_config_btn, self.colors['secondary'], '#7f8c8d')
        
    def _add_placeholder(self, entry: tk.Entry, var: tk.StringVar, placeholder: str) -> None:
        """为输入框添加占位符文字
        
        Args:
            entry: 输入框组件
            var: 关联的StringVar变量
            placeholder: 占位符文字
        """
        # 初始显示占位符
        var.set(placeholder)
        entry.config(fg='#999999')
        
        def on_focus_in(event):
            if var.get() == placeholder:
                var.set('')
                entry.config(fg=self.colors['fg'])
        
        def on_focus_out(event):
            if not var.get():
                var.set(placeholder)
                entry.config(fg='#999999')
        
        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)
    
    def _add_hover_effect(self, button: tk.Button, normal_color: str, hover_color: str) -> None:
        """为按钮添加悬停效果
        
        Args:
            button: 按钮组件
            normal_color: 正常颜色
            hover_color: 悬停颜色
        """
        def on_enter(e):
            button['bg'] = hover_color
        
        def on_leave(e):
            button['bg'] = normal_color
        
        button.bind('<Enter>', on_enter)
        button.bind('<Leave>', on_leave)
    
    def setup_logging(self) -> None:
        """设置日志系统"""
        # 为单文件处理设置日志处理器
        gui_handler = GUILogHandler(self.log_text)
        gui_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
        gui_handler.setLevel(logging.INFO)
        
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.handlers.clear()
        logger.addHandler(gui_handler)
        
    def browse_file(self) -> None:
        """浏览文件对话框"""
        file_path = filedialog.askopenfilename(
            title="选择DOCX文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir_var.set(dir_path)
            self.output_entry.config(fg=self.colors['fg'])  # 设置为正常颜色
            
    def browse_batch_input(self):
        """浏览批量输入目录"""
        dir_path = filedialog.askdirectory(title="选择输入目录")
        if dir_path:
            self.batch_input_var.set(dir_path)
            
    def browse_batch_output(self):
        """浏览批量输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.batch_output_var.set(dir_path)
            self.batch_output_entry.config(fg=self.colors['fg'])  # 设置为正常颜色
            
    def process_file(self) -> None:
        """处理单个文件"""
        if self.processing:
            return
            
        file_path = self.file_path_var.get().strip()
        if not file_path:
            messagebox.showerror("错误", "请选择一个DOCX文件")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在")
            return
            
        if not file_path.lower().endswith('.docx'):
            messagebox.showerror("错误", "请选择DOCX格式的文件")
            return
            
        # 清空日志区域
        self.log_text.delete(1.0, tk.END)
        
        # 在新线程中处理文件
        threading.Thread(target=self._process_file_thread, args=(file_path,), daemon=True).start()
        
    def _process_file_thread(self, file_path):
        """在后台线程中处理文件"""
        try:
            self.processing = True
            self.status_var.set("处理中...")
            self.process_btn.config(state=tk.DISABLED)
            
            # 应用设置到配置
            self._apply_settings_to_config()
            
            extractor = DocumentExtractor()
            content = extractor.extract_content(file_path)
            
            # 确定输出路径
            output_dir = self.output_dir_var.get().strip()
            # 检查是否为占位符文字或空
            if output_dir and not output_dir.startswith('（可选）'):
                output_path = Path(output_dir) / f"{Path(file_path).stem}.md"
            else:
                output_path = Path(file_path).with_suffix('.md')
            
            # 确保输出目录存在
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 保存文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            self.last_output_file = str(output_path)
            
            # 更新UI
            self.root.after(0, self._on_process_complete, str(output_path))
            
        except Exception as e:
            self.root.after(0, self._on_process_error, str(e))
        finally:
            self.processing = False
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            
    def _on_process_complete(self, output_path):
        """处理完成回调"""
        self.status_var.set("处理完成")
        self.open_output_btn.config(state=tk.NORMAL)
        messagebox.showinfo("成功", f"文件处理完成!\n输出文件: {output_path}")
        
    def _on_process_error(self, error_msg):
        """处理错误回调"""
        self.status_var.set("处理失败")
        messagebox.showerror("错误", f"处理文件时发生错误: {error_msg}")
        
    def process_batch(self):
        """批量处理文件"""
        if self.processing:
            return
            
        input_dir = self.batch_input_var.get().strip()
        if not input_dir:
            messagebox.showerror("错误", "请选择输入目录")
            return
            
        if not os.path.exists(input_dir):
            messagebox.showerror("错误", "输入目录不存在")
            return
            
        output_dir = self.batch_output_var.get().strip()
        # 检查是否为占位符文字或空，如果是则使用输入目录
        if not output_dir or output_dir.startswith('（可选）'):
            output_dir = input_dir
            
        # 清空日志
        self.batch_log_text.delete(1.0, tk.END)
        
        # 在新线程中处理
        threading.Thread(target=self._process_batch_thread, args=(input_dir, output_dir), daemon=True).start()
        
    def _process_batch_thread(self, input_dir, output_dir):
        """批量处理后台线程"""
        try:
            self.processing = True
            self.root.after(0, lambda: self.batch_process_btn.config(state=tk.DISABLED))
            self.root.after(0, lambda: self.status_var.set("查找文件中..."))
            
            # 查找DOCX文件
            docx_files = []
            input_path = Path(input_dir)
            
            if self.recursive_var.get():
                pattern = '**/*.docx'
            else:
                pattern = '*.docx'
                
            for file_path in input_path.glob(pattern):
                if not file_path.name.startswith('~$'):
                    docx_files.append(file_path)
            
            if not docx_files:
                self.root.after(0, lambda: messagebox.showwarning("警告", "未找到DOCX文件"))
                return
            
            # 更新进度条
            total_files = len(docx_files)
            self.root.after(0, lambda: self.progress_bar.config(maximum=total_files))
            self.root.after(0, lambda: self.progress_var.set(f"0/{total_files}"))
            
            # 处理文件
            success_count = 0
            extractor = DocumentExtractor()
            
            for i, file_path in enumerate(docx_files):
                try:
                    self.root.after(0, lambda: self.status_var.set(f"处理中: {file_path.name}"))
                    
                    content = extractor.extract_content(str(file_path))
                    
                    # 确定输出路径
                    relative_path = file_path.relative_to(input_path)
                    output_path = Path(output_dir) / relative_path.with_suffix('.md')
                    
                    # 确保输出目录存在
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    
                    # 保存文件
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    
                    success_count += 1
                    
                    # 更新进度
                    progress = i + 1
                    self.root.after(0, lambda p=progress: self.progress_bar.config(value=p))
                    self.root.after(0, lambda p=progress, t=total_files: self.progress_var.set(f"{p}/{t}"))
                    
                    # 添加日志
                    log_msg = f"✓ {file_path.name} -> {output_path.name}\n"
                    self.root.after(0, lambda msg=log_msg: self.batch_log_text.insert(tk.END, msg))
                    self.root.after(0, lambda: self.batch_log_text.see(tk.END))
                    
                except Exception as e:
                    log_msg = f"✗ {file_path.name}: {str(e)}\n"
                    self.root.after(0, lambda msg=log_msg: self.batch_log_text.insert(tk.END, msg))
                    self.root.after(0, lambda: self.batch_log_text.see(tk.END))
            
            # 完成处理
            self.batch_output_dir = output_dir
            self.root.after(0, lambda: self._on_batch_complete(success_count, total_files))
            
        except Exception as e:
            self.root.after(0, lambda: self._on_batch_error(str(e)))
        finally:
            self.processing = False
            self.root.after(0, lambda: self.batch_process_btn.config(state=tk.NORMAL))
            
    def _on_batch_complete(self, success_count, total_files):
        """批量处理完成回调"""
        self.status_var.set(f"批量处理完成: {success_count}/{total_files}")
        self.open_batch_output_btn.config(state=tk.NORMAL)
        messagebox.showinfo("完成", f"批量处理完成!\n成功: {success_count}/{total_files}")
        
    def _on_batch_error(self, error_msg):
        """批量处理错误回调"""
        self.status_var.set("批量处理失败")
        messagebox.showerror("错误", f"批量处理时发生错误: {error_msg}")
        
    def open_output_file(self):
        """打开输出文件"""
        if hasattr(self, 'last_output_file') and os.path.exists(self.last_output_file):
            os.startfile(self.last_output_file)
        
    def open_batch_output(self):
        """打开批量输出目录"""
        if hasattr(self, 'batch_output_dir') and os.path.exists(self.batch_output_dir):
            os.startfile(self.batch_output_dir)
            
    def _apply_settings_to_config(self):
        """应用界面设置到配置对象"""
        self.config.text_width = self.width_var.get()
        self.config.text_indent = self.indent_var.get()
        self.config.base_column_width = self.col_width_var.get()
        
    def save_config(self):
        """保存配置"""
        try:
            self._apply_settings_to_config()
            config_path = filedialog.asksaveasfilename(
                title="保存配置文件",
                defaultextension=".yaml",
                filetypes=[
                    ("YAML文件", "*.yaml"), 
                    ("YAML文件", "*.yml"),
                    ("JSON文件", "*.json"), 
                    ("所有文件", "*.*")
                ]
            )
            if config_path:
                self.config.save_to_file(config_path)
                messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            
    def load_config(self):
        """加载配置"""
        try:
            config_path = filedialog.askopenfilename(
                title="加载配置文件",
                filetypes=[
                    ("YAML文件", "*.yaml"), 
                    ("YAML文件", "*.yml"),
                    ("JSON文件", "*.json"), 
                    ("所有文件", "*.*")
                ]
            )
            if config_path:
                self.config.load_from_file(config_path)
                self._update_ui_from_config()
                messagebox.showinfo("成功", "配置已加载")
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {str(e)}")
            
    def reset_config(self):
        """重置配置为默认值"""
        self.config = Config()
        self._update_ui_from_config()
        messagebox.showinfo("成功", "配置已重置为默认值")
        
    def _update_ui_from_config(self):
        """从配置更新界面"""
        self.width_var.set(self.config.text_width)
        self.indent_var.set(self.config.text_indent)
        self.col_width_var.set(self.config.base_column_width)
        
    def run(self) -> None:
        """运行GUI"""
        self.root.mainloop()


def main() -> int:
    """GUI主函数
    
    Returns:
        退出码，0表示成功，非0表示失败
    """
    try:
        app = DocumentExtractorGUI()
        app.run()
    except Exception as e:
        print(f"启动GUI失败: {str(e)}")
        return 1
    return 0


if __name__ == '__main__':
    import sys
    sys.exit(main())
