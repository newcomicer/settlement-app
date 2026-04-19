# settlement-app — 路跑賽事經費結算系統

## 這個工具是做什麼的
iRunner 匯出的報名 Excel → 自動解析組別人數、財務數字 → 前端確認 / 調整 → 匯出 PDF 費用申請單

使用者是 Johnny（路跑賽事規劃與計時服務業務主管），這是他日常結算賽事費用的內部工具。

## 技術架構

| 層 | 技術 |
|---|---|
| 後端 | Python + Flask，port 5001 |
| 前端 | `templates/index.html`（單頁，無前端框架） |
| Excel 解析 | openpyxl，讀「報名資料」工作表 |
| PDF 產生 | Playwright（headless Chromium 渲染 `templates/pdf_pages.html`） |

## 啟動方式
```bash
cd tools/settlement-app
python app.py
# 開啟 http://127.0.0.1:5001
```
或直接雙擊 `啟動.command`（macOS）。

## 核心邏輯（app.py）

### parse_excel()
- 找「報名資料」工作表，動態偵測欄位名稱
- 優先讀合計列（最後一個有實繳金額的空日期列），否則逐訂單加總
- 公關人數：優先讀「免費名單」工作表，否則從訂單類型判斷
- 加購數量：動態偵測含「加購」或「加價購」的欄位
- 組別單價：從「報名項目費用」欄位自動推算最常見金額

### calculate_settlement()
- **方式一**：實繳金額 - ATM/超商手續費 - 退費 - 晶片押金 - 計時費
- **方式二**：報名費 + 加購費 + 郵寄費 + 溢收 - 計時費
- 計時費支援百分比項目（is_percent: true）
- 兩方式若相符代表帳務正確

## 檔案結構
```
settlement-app/
├── app.py              # Flask 主程式（解析 + 計算 + PDF）
├── requirements.txt    # flask, openpyxl, playwright
├── 啟動.command        # macOS 一鍵啟動
├── templates/
│   ├── index.html      # 主介面（上傳 Excel、填手動欄位、下載 PDF）
│   └── pdf_pages.html  # PDF 版面（Playwright 渲染用）
└── static/
    └── uploads/        # 暫存上傳的 Excel（不 commit）
```

## 改版注意事項
- `static/uploads/` 不要 commit（已在 .gitignore）
- 新增欄位偵測邏輯時，要在 `parse_excel()` 的欄位動態偵測部分處理，不要寫死欄位索引
- PDF 樣式修改請動 `templates/pdf_pages.html`，A4 頁面、15mm 四邊邊距
- 測試用 Excel 請使用真實 iRunner 匯出格式（含「報名資料」工作表）

## GitHub
repo：https://github.com/newcomicer/settlement-app（公開）
