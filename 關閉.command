#!/bin/bash
echo "🛑 關閉經費結算系統..."

PIDS=$(lsof -ti :5001)

if [ -z "$PIDS" ]; then
    echo "✅ Port 5001 沒有在跑的程序，不需要關閉。"
else
    echo "找到程序 PID：$PIDS，正在關閉..."
    kill $PIDS
    sleep 1
    echo "✅ 已關閉。"
fi
