# CAD文本提取与翻译工具

一个专业的CAD文本处理工具，支持DXF文件的中英文文本提取、翻译和回填功能。

## 项目结构

```
cad-text-translator/
├── gui.py                    # GUI主界面
├── 回填.py                   # 文本回填功能
├── extract_texts.py          # 文本提取功能
├── font_config.py           # 字体配置管理
├── logger_config.py         # 日志配置
├── dxf_cleaner.py          # DXF文件清理工具
├── dxf_text_extractor.py   # DXF文本提取器
├── haochen_optimized_converter.py # 优化转换器
├── 命令行专用/              # 命令行工具
│   ├── main_processor.py    # 主处理器
│   ├── simple_processor.py  # 简单处理器
│   └── 提取.py             # 命令行提取工具
├── docs/                   # 详细文档目录（包含分析报告）
├── requirements.txt        # 依赖包列表
├── start_gui.bat          # Windows启动脚本
└── start_gui.ps1          # PowerShell启动脚本
```

## 主要功能

### GUI专用模块
- **gui.py**: 图形用户界面，提供直观的操作体验
- **回填.py**: 支持智能翻译的文本回填功能
- **extract_texts.py**: 文本提取与处理

### 命令行专用模块
- **main_processor.py**: 完整的命令行处理流程
- **simple_processor.py**: 简化的处理流程
- **提取.py**: 专门的文本提取工具

### 共享模块
- **font_config.py**: 字体配置管理，支持中英文字体设置
- **logger_config.py**: 统一的日志管理
- **dxf_cleaner.py**: DXF文件清理和优化
- **dxf_text_extractor.py**: 核心文本提取引擎
- **haochen_optimized_converter.py**: 优化的转换算法

## 快速开始

### 环境要求
- Python 3.8+
- 依赖包见 requirements.txt

### 安装依赖
```bash
pip install -r requirements.txt
```

### 启动GUI版本
```bash
# Windows
start_gui.bat

# 或使用PowerShell
.\start_gui.ps1

# 或直接运行Python
python gui.py
```

### 命令行使用
```bash
# 使用主处理器
python 命令行专用/main_processor.py input.dxf

# 使用简单处理器
python 命令行专用/simple_processor.py input.dxf

# 仅提取文本
python 命令行专用/提取.py input.dxf
```

## 核心流程

1. **文本提取**: 从DXF文件中提取所有文本实体
2. **智能翻译**: 支持中英文互译，包含空翻译处理优化
3. **文本回填**: 将翻译后的文本回填到DXF文件
4. **字体配置**: 自动配置合适的中英文字体
5. **文件清理**: 优化DXF文件结构

## 开发者指南

### 架构设计
项目采用模块化设计，分为以下层次：

#### 表现层
- **GUI模块**: gui.py - 图形界面
- **CLI模块**: 命令行专用文件夹 - 命令行接口

#### 业务层
- **提取服务**: extract_texts.py, dxf_text_extractor.py
- **翻译服务**: 回填.py (包含智能翻译功能)
- **转换服务**: haochen_optimized_converter.py

#### 工具层
- **配置管理**: logger_config.py, font_config.py, haochen_optimized_converter.py, dxf_cleaner.py
- **文件处理**: dxf_cleaner.py

### 扩展开发
1. 新增转换器：继承基础转换器类
2. 自定义字体：修改 font_config.py
3. 添加翻译引擎：扩展翻译服务模块

## 文档

详细的技术文档和分析报告请查看 `docs/` 目录：
- 使用说明.md
- 性能优化分析报告.md
- 智能翻译功能报告.md
- DXF格式分析总结.md
- 等等...

## 许可证

本项目采用 MIT 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。
