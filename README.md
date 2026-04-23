# ⚡ 經費結算系統

iRunner Excel 匯入 → 自動計算結算金額 → 匯出 PDF 費用申請單

---

## 安裝教學（第一次才需要做）

### Step 1：安裝 Python 3

打開瀏覽器，前往 👉 https://www.python.org/downloads/

點擊黃色的 **Download Python 3.x.x** 按鈕，下載後照步驟安裝。

> ⚠️ Windows 用戶：安裝時第一個畫面最下方有個 **「Add python.exe to PATH」，一定要勾起來**，否則終端機找不到 Python。

安裝完成後，打開終端機確認：

**Mac 用戶：**
```
python3 --version
```

**Windows 用戶：**
```
python --version
```

有出現版本號就 OK（例如 `Python 3.12.0`）

---

### Step 2：下載程式

**方式 A：直接下載 ZIP（建議，不需要 Git）**

前往 👉 https://github.com/newcomicer/settlement-app

點右上角綠色 **Code** 按鈕 → **Download ZIP** → 解壓縮到桌面

> 解壓後資料夾名稱會是 `settlement-app-main`

**方式 B：用 Git clone（適合需要後續更新的用戶）**

先確認有安裝 Git：

- **Mac**：終端機輸入 `git --version`，有版本號即已安裝；若沒有，前往 👉 https://git-scm.com/download/mac
- **Windows**：前往 👉 https://git-scm.com/download/win 下載安裝，過程保持預設選項即可

安裝完成後執行：

```
git clone https://github.com/newcomicer/settlement-app
```

> 往後只需在資料夾內執行 `git pull` 即可取得最新版本

---

### Step 3：安裝所需套件

打開終端機，進入資料夾：

**Mac 用戶：**
```bash
cd ~/Desktop/settlement-app-main
```

**Windows 用戶：**
```
cd C:\Users\你的名字\Desktop\settlement-app-main
```

> 若是用 Git clone 下載，資料夾名稱為 `settlement-app`（沒有 `-main`）

然後執行以下兩行（逐行貼上，等每行跑完再貼下一行）：

**Mac 用戶：**
```bash
pip3 install -r requirements.txt
playwright install chromium
```

**Windows 用戶：**
```
pip install -r requirements.txt
python -m playwright install chromium
```

> 第二行會下載約 150MB，這是產生 PDF 需要用到的瀏覽器核心，請耐心等候。  
> 出現 `Chromium 版本號 downloaded` 之類的訊息就代表成功了。

---

## 每次使用

### Mac 用戶
- **啟動**：雙擊資料夾裡的 **「啟動.command」**，程式會自動開啟瀏覽器
- **關閉**：雙擊資料夾裡的 **「關閉.command」**

> 第一次執行可能會出現安全性警告，請前往  
> **系統設定 → 隱私權與安全性** → 點「仍要開啟」  
> 啟動和關閉的 .command 各需允許一次

### Windows 用戶
雙擊資料夾裡的 **「啟動.bat」**，程式會自動開啟瀏覽器。

關閉跳出的黑色視窗即可停止程式。

> 如果出現「找不到 flask」的錯誤，請先完成 Step 3 的套件安裝再試一次。

---

## 使用方式

1. 從 iRunner 後台匯出報名資料 Excel（`.xlsx`）
2. 將 Excel 拖拉到左側上傳區，系統自動解析並帶入數值
3. 若已有設定檔，點左側「**載入設定檔**」自動帶入匯款帳戶與服務費項目
4. 確認並補齊：活動名稱、請款日期、匯款帳戶
5. 調整計費單價、公關人數、服務費項目（首次填寫完可點「**儲存設定檔**」留存，下次直接載入）
6. 右側即時預覽結算結果
7. 點右上角「**匯出 PDF**」下載費用申請單

---

## 常見問題

**Q：執行後瀏覽器沒有自動打開？**  
手動打開瀏覽器，輸入 http://127.0.0.1:5001

**Q：匯出 PDF 失敗？**  
請確認有執行過 Step 3 的 `playwright install chromium`

**Q：Excel 解析失敗？**  
請確認是從 iRunner 後台匯出的「報名資料」格式，且副檔名為 `.xlsx`

**Q：要關閉程式？**  
- **Mac**：雙擊資料夾裡的「關閉.command」
- **Windows**：關閉啟動時跳出的黑色視窗
