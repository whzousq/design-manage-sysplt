#!/bin/bash
# 设计项目进度管理系统 - 启动脚本

echo "========================================"
echo "  设计项目进度管理系统 - 环境检查"
echo "========================================"

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python3，请先安装 Python 3.8+"
    exit 1
fi

echo "Python 版本: $(python3 --version)"

# Create virtual environment if not exists
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "激活虚拟环境..."
# shellcheck source=/dev/null
source venv/bin/activate

# Install dependencies
echo "安装依赖..."
pip install -r requirements.txt -q

echo ""
echo "========================================"
echo "  启动服务器..."
echo "========================================"
echo ""
echo "  本地访问: http://localhost:5000"
echo "  局域网访问: http://<本机IP>:5000"
echo ""
echo "  默认账号: admin / admin123"
echo "  按 Ctrl+C 停止服务器"
echo ""
echo "========================================"

# Start server
python app.py
