#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"
echo "正在启动 绍兴照片智能 AI 质检系统..."
python3 app.py
EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "程序异常退出，退出码: $EXIT_CODE"
    read -n 1 -s -r -p "按任意键退出窗口..."
fi
