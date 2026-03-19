@echo off
chcp 65001 >nul
echo ============================================
echo   Simulink 文档生成器 - Windows 安装脚本
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
pip install python-docx lxml

echo.
echo [2/3] 安装打包工具...
pip install pyinstaller

echo.
echo [3/3] 打包成 exe...
pyinstaller --onefile --windowed --name "Simulink文档生成器" --clean main.py

echo.
echo ============================================
echo   打包完成！
echo   可执行文件: dist\Simulink文档生成器.exe
echo ============================================
pause