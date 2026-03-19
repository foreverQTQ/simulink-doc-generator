# Simulink 文档生成器

从 MATLAB Simulink 模型文件自动生成 Word 设计文档。

## ✨ 功能特点

- ✅ **纯 Python 实现** - 无需 MATLAB 许可证
- ✅ **智能功能描述** - 自动推断子系统功能
- ✅ **美观文档格式** - 表格化展示、专业排版
- ✅ **双模式运行** - GUI 图形界面 + 命令行
- ✅ **可打包成 exe** - 无需 Python 环境也可使用

## 📦 快速开始

### Windows 用户

1. 双击 `build_windows.bat` 打包成 exe
2. 运行 `dist\Simulink文档生成器.exe`

### 有 Python 的用户

```bash
# 安装依赖
pip install python-docx lxml

# GUI 模式
python main.py

# 命令行模式
python main.py model.slx
```

## 📖 详细说明

请查看 [使用指南.md](使用指南.md)

## 📁 项目结构

```
simulink-doc-generator/
├── main.py              # 主入口
├── requirements.txt     # 依赖清单
├── build_windows.bat    # Windows 打包脚本
├── build_macos.sh       # macOS 打包脚本
├── core/                # 核心模块
│   ├── parser.py        # SLX 解析器
│   ├── analyzer.py      # 模型分析器
│   └── generator.py     # 文档生成器
└── gui/                 # GUI 界面
    └── main_window.py
```

## 🎯 测试结果

已在以下模型上测试通过：
- K4BEV_Model (2125模块, 193子系统, 7层级)

---

Made with ❤️ by 文婧 🐰