# DOCX文档提取器

专业的Word文档内容提取工具，将DOCX格式文档转换为Markdown格式，完美保持原文档结构和表格布局。

## 特性

- 🎨 **现代化GUI** - 简洁美观的图形界面，操作直观
- 📋 **表格处理** - 智能处理复杂表格，支持合并单元格
- 📁 **批量处理** - 支持目录递归搜索和批量转换
- ⚙️ **配置管理** - 可视化配置，支持保存/加载
- 🌐 **中文优化** - 完美支持中文字符宽度计算

## 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 启动程序

```bash
# 启动GUI界面
python src/main.py
```

或直接双击运行 `src/main.py`

## 项目结构

```text
docs2txt/
├── src/                    # 源代码
│   ├── main.py            # 主入口
│   ├── core.py            # 核心提取逻辑
│   ├── gui.py             # 图形界面
│   ├── config.py          # 配置管理
│   └── docx_extractor.py  # 文档提取器
├── config/                # 配置文件
├── samples/               # 示例文档
├── output/                # 输出目录
└── dist/                  # 可执行文件
```

## 配置说明

编辑 `config/default.yaml` 自定义处理参数：

```yaml
# 文本格式配置
text_width: 80              # 文本行宽度
text_indent: "    "         # 段落缩进

# 表格格式配置  
base_column_width: 15       # 基础列宽
level_2_multiplier: 2       # 中等列宽倍数
```

## 构建可执行文件

```bash
# Windows
build_exe.bat

# 生成 dist/docx_extractor.exe
```

## 许可证

MIT License
