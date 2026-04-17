#!/bin/bash
cd "$(dirname "$0")"
echo "🚀 啟動經費結算系統..."
python3 app.py &
sleep 2
open http://127.0.0.1:5001
wait
