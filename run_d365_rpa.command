#!/bin/bash
# 雙擊此檔案即可自動開啟 Terminal，啟用 venv 並執行 run_d365_rpa.py

# 切換到本檔案所在的目錄（無論從哪裡雙擊都能正確定位專案路徑）
cd "$(dirname "$0")"

echo "=== 啟用虛擬環境 (venv) ==="
source venv/bin/activate

echo "=== 開始執行 run_d365_rpa.py ==="
python3 run_d365_rpa.py

echo ""
echo "=== 程式執行完畢 ==="
echo "按下 Enter 鍵關閉視窗..."
read
