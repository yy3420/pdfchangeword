@echo off
echo 安装PDF转Word软件所需依赖...

REM 检查Python是否已安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python未安装，请先安装Python 3.8或更高版本
    echo 可以从 https://www.python.org/downloads/ 下载安装
    pause
    exit /b 1
)

REM 创建Python虚拟环境
echo 创建Python虚拟环境...
python -m venv venv
call venv\Scripts\activate.bat

REM 安装Python依赖
echo 安装Python依赖...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo ==========================================================
echo 注意：您还需要手动安装Tesseract OCR才能使用文字识别功能
echo 1. 从 https://github.com/UB-Mannheim/tesseract/wiki 下载并安装Tesseract
echo 2. 确保将Tesseract添加到系统环境变量PATH中
echo ==========================================================

echo 依赖安装完成！
echo 使用以下命令启动软件：
echo venv\Scripts\activate.bat ^& python run.py

pause 