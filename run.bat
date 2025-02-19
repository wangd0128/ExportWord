@echo off
echo 正在启动Word表格处理工具...
echo 请确保已经安装了所有必要的Python包

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo Python未安装，请先安装Python！
    pause
    exit /b
)

:: 检查必要的包是否安装
echo 检查必要的包...
python check_requirements.py
if errorlevel 1 (
    echo 正在安装必要的包...
    pip install gradio pandas python-docx pywin32
)

:: 创建outputs目录（如果不存在）
if not exist "outputs" mkdir outputs

:: 运行主程序
python main.py

:: 如果程序异常退出，暂停显示错误信息
if errorlevel 1 (
    echo 程序运行出错！
    pause
) 