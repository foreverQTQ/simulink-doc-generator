#!/bin/bash
# Simulink 文档生成器 - macOS/Linux 打包脚本

echo "============================================"
echo "  Simulink 文档生成器 - 打包脚本"
echo "============================================"
echo

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未找到 Python3"
    exit 1
fi

echo "[1/3] 安装依赖..."
pip3 install --break-system-packages python-docx lxml

echo
echo "[2/3] 安装打包工具..."
pip3 install --break-system-packages pyinstaller

echo
echo "[3/3] 打包成可执行文件..."
pyinstaller --onefile --windowed --name "Simulink文档生成器" --clean main.py

echo
echo "============================================"
echo "  打包完成！"
echo "  可执行文件: dist/Simulink文档生成器"
echo "============================================"