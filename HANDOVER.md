# ASUS News Intelligence Dashboard — 交接文件 v3

> 本文件描述目前系統的完整架構、設計規範、Google 快訊整合與自動化 GitHub 發布維護方式，供 AI 助理每週更新使用。

---

## 一、系統概覽

每週自動抓取 PC 產業與國際經濟相關新聞，透過 Gemini API 進行 AI 摘要，寫入 Excel，並產生 HTML 週報看板，同時**自動發布至 GitHub Pages 線上網站**。

```
news_collector.py     ← 主程式（抓新聞 + AI摘要 + 寫Excel + 產HTML + 自動推送到 GitHub）
generate_html_only.py ← 只重新產HTML，不抓新聞（讀現有Excel + 自動推送到 GitHub）
.env                  ← 存放 API Key、DIGITIMES 帳密與 Google 快訊 RSS 網址
.gitignore            ← 自動資安防護（禁止上傳 .env, News.xlsx, __pycache__）
index.html            ← 最新週報（永遠指向當週，自動上傳）
2026wkXX.html         ← 歷史週報歸檔（自動上傳）
D:\ASUS\News & Report\News\News.xlsx  ← 主要資料來源
```

---

## 二、執行方式

在 CMD (命令提示字元) 中執行以下指令：

```cmd
cd /d D:\ASUS\Anti-NotebookLM\NEWS

:: 每週抓新聞（自動由快訊/RSS抓取，AI摘要並推送到 GitHub Pages）
python news_collector.py   (或執行 run_news_collector.bat)

:: 只重新產HTML（修改 Excel 後，不抓新聞直接更新網路上的 Dashboard）
python generate_html_only.py  (或執行 run_html_only.bat)
```

---

## 三、Google 快訊 RSS 整合 (國際經濟)

為了提升「國際經濟」新聞品質，系統已整合 Google 快訊 (Google Alerts) RSS 訂閱源：

1. **環境變數設定 (.env)**：
   ```env
   GOOGLE_ALERT_RSS_ECON=https://www.google.com.tw/alerts/feeds/xxxxxxxxxx/xxxxxxxxxx
   ```
2. **優先級與排除排序演算法**：
   - 頂級優先：標題含 `中國`、`美國`、`美中`、`聯準會`、`Fed`、`通膨`、`降息`
   - 宏觀優先：標題含 `全球經濟`、`全球市場`、`國際經濟`、`全球成長`
   - 自動降級/排除：標題含 `董座`、`董事長`、`台股`、`台廠`、`三陽`、`光陽`、`台積電` 等本土企業財報新聞降至最低優先度。

---

## 四、一鍵自動發布至 GitHub Pages (Auto Deploy)

程式內建 `auto_deploy_to_github()` 函式，執行流程如下：

1. **自動搜尋 Git 根目錄**：自動定位包含 `.git` 的工作目錄。
2. **自動 Git 命令**：自動執行 `git add .` -> `git commit -m "Auto update..."` -> `git push`。
3. **備援對齊機制**：若遇首次分支推送到 `origin main`，自動執行 `git push -u origin main` 與 `--force` 強制對齊。
4. **編碼安全**：背景指令一律採用 `UTF-8` 搭配 `errors='replace'` 讀取，防止 Windows `cp950` 解碼例外。
5. **發布確認**：推送成功後，GitHub Pages 約在 30 秒至 2 分鐘內完成線上部署。線上讀者可透過 `https://AngelaHNWang.github.io/Weekly-News-Summary/` 查看最新看板。

---

## 五、重要設定與檔案規範

### `.env` 檔案（資安重地，絕對不外洩）
```env
GEMINI_API_KEY=...
DIGITIMES_USER=...
DIGITIMES_PASSWORD=...
GOOGLE_ALERT_RSS_ECON=...
```

### Excel 路徑
```python
EXCEL_PATH = r"D:\ASUS\News & Report\News\News.xlsx"
```

---

## 六、HTML 設計規範與配色

| 分類 | `--cat-color` | `--cat-bg` | `--cat-border` |
|------|--------------|------------|----------------|
| 國際經濟 | `#b45309` | `rgba(180,83,9,0.06)` | `rgba(180,83,9,0.2)` |
| PC產業 | `#0369a1` | `rgba(3,105,161,0.06)` | `rgba(3,105,161,0.2)` |
| 上游廠商動態 | `#6d28d9` | `rgba(109,40,217,0.06)` | `rgba(109,40,217,0.2)` |
| 各品牌廠動態 | `#047857` | `rgba(4,120,87,0.06)` | `rgba(4,120,87,0.2)` |

---

## 七、注意事項與維護

1. 執行前請確認 **Excel 已關閉**，否則 `openpyxl` 無法寫入。
2. `.gitignore` 已設定完成，`.env` 與 `News.xlsx` **絕對不會**上傳到 GitHub。
3. 若線上網址未及時更換，請在瀏覽器按 **`Ctrl + F5`** 清除快取重整即可。
