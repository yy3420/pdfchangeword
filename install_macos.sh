#!/bin/bash
# MacOS安装脚本

echo "安装PDF转Word软件所需依赖..."

# 检查并安装Homebrew
if ! command -v brew &> /dev/null; then
    echo "正在安装Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi

# 安装Tesseract OCR引擎
echo "安装Tesseract OCR引擎..."
brew install tesseract
brew install tesseract-lang  # 安装语言包

# 安装Poppler (pdf2image依赖)
echo "安装Poppler (用于PDF图像转换)..."
brew install poppler

# 创建Python虚拟环境
echo "创建Python虚拟环境..."
python3 -m venv venv
source venv/bin/activate

# 安装Python依赖
echo "安装Python依赖..."
pip install --upgrade pip
pip install -r requirements.txt

echo "依赖安装完成！"
echo "使用以下命令启动软件："
echo "source venv/bin/activate && python run.py" 