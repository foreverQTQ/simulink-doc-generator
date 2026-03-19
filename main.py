#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simulink 文档生成器
从 MATLAB Simulink 模型文件生成 Word 设计文档

用法:
    python main.py              # 启动 GUI
    python main.py model.slx    # 命令行模式
"""

import sys
import os
from pathlib import Path

# 添加当前目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from gui.main_window import main as gui_main
from core.parser import parse_slx
from core.analyzer import ModelAnalyzer
from core.generator import DocGenerator


def cli_mode(slx_path: str, output_path: str = None):
    """命令行模式"""
    print(f"Simulink 文档生成器 v1.0")
    print(f"=" * 50)
    
    # 检查文件
    if not os.path.exists(slx_path):
        print(f"错误: 文件不存在 - {slx_path}")
        return 1
    
    # 设置输出路径
    if output_path is None:
        base_name = Path(slx_path).stem
        output_dir = os.path.dirname(slx_path)
        output_path = os.path.join(output_dir, f"{base_name}_设计文档.docx")
    
    try:
        # 解析
        print(f"\n[1/3] 正在解析模型...")
        model = parse_slx(slx_path)
        print(f"  模型名称: {model.name}")
        
        # 分析
        print(f"\n[2/3] 正在分析模型...")
        analyzer = ModelAnalyzer(model)
        overview = analyzer.get_model_overview()
        print(f"  模块总数: {overview['total_blocks']}")
        print(f"  子系统数量: {overview['total_subsystems']}")
        
        # 生成
        print(f"\n[3/3] 正在生成文档...")
        generator = DocGenerator(model, analyzer)
        generator.generate(output_path)
        print(f"  输出文件: {output_path}")
        
        print(f"\n✓ 文档生成成功！")
        return 0
        
    except Exception as e:
        print(f"\n✗ 错误: {e}")
        import traceback
        traceback.print_exc()
        return 1


def main():
    """主函数"""
    if len(sys.argv) > 1:
        # 命令行模式
        slx_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        return cli_mode(slx_path, output_path)
    else:
        # GUI 模式
        gui_main()
        return 0


if __name__ == "__main__":
    sys.exit(main())