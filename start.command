#!/bin/zsh

cd "$(dirname "$0")"

echo "正在启动 Jack 市场沟通助手..."
echo "如果浏览器没有自动打开，请手动访问：http://localhost:8505"
echo ""

lsof -ti tcp:8505 | xargs kill -9 2>/dev/null

.venv/bin/streamlit run app.py --server.port 8505 --server.address localhost
