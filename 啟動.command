#!/bin/bash
cd "$(dirname "$0")"
echo "🚀 啟動經費結算系統..."

# 找到有安裝 flask 的 python3
PYTHON=""
for p in \
  "/Library/Developer/CommandLineTools/Library/Frameworks/Python3.framework/Versions/3.9/bin/python3.9" \
  "/opt/homebrew/bin/python3" \
  "python3" \
  "python"; do
  if $p -c "import flask" 2>/dev/null; then
    PYTHON=$p
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "❌ 找不到已安裝 flask 的 Python，請執行："
  echo "   pip3 install -r requirements.txt"
  read -p "按 Enter 關閉..."
  exit 1
fi

echo "✅ 使用 Python：$PYTHON"
$PYTHON app.py &
sleep 2
open http://127.0.0.1:5001
wait
