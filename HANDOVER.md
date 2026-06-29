# ASUS News Intelligence Dashboard — 交接文件 v2

> 本文件描述目前系統的完整架構、設計規範與維護方式，供 AI 助理（如 Gemini）每週更新使用。

---

## 一、系統概覽

每週自動抓取 PC 產業相關新聞，透過 Gemini API 進行 AI 摘要，寫入 Excel，並產生 HTML 週報看板。

```
news_collector.py   ← 主程式（抓新聞 + AI摘要 + 寫Excel + 產HTML）
generate_html_only.py  ← 只重新產HTML，不抓新聞（讀現有Excel）
.env                ← 存放 API Key 與帳號密碼
index.html          ← 最新週報（永遠指向當週）
2026wkXX.html       ← 歷史週報歸檔
D:\ASUS\News & Report\News\News.xlsx  ← 主要資料來源
```

---

## 二、執行方式

```cmd
cd /d D:\ASUS\Anti-NotebookLM\NEWS

:: 每週抓新聞（會詢問確認週次）
python news_collector.py

:: 只重新產HTML（不抓新聞，直接讀Excel）
python generate_html_only.py
```

執行 `news_collector.py` 時，系統會：
1. 讀取 Excel 最新週次，自動建議本次週次（max週+1）
2. **互動詢問確認**：直接 Enter 接受建議值，或輸入數字（如 `27`）覆蓋
3. 抓取新聞、AI摘要、寫入 Excel
4. 自動產生 `index.html` 與 `{year}wk{XX}.html` 兩個檔案

---

## 三、重要設定

### `.env` 檔案（不可外洩）
```
GEMINI_API_KEY=...
DIGITIMES_USER=...
DIGITIMES_PASSWORD=...
```

### Excel 路徑
```python
EXCEL_PATH = r"D:\ASUS\News & Report\News\News.xlsx"
```
注意：路徑含 `&` 符號，使用 raw string。

### 新聞分類與搜尋關鍵字（`news_collector.py` 頂部 `CATEGORIES` dict）
```python
CATEGORIES = {
    "國際經濟":    "全球經濟 科技產業",
    "PC產業":     "PC 筆電 筆記型電腦",
    "上游廠商動態": "半導體 供應鏈 晶片",
    "各品牌廠動態": "ASUS 宏碁 聯想 惠普 蘋果 筆電",
}
```

---

## 四、HTML 設計規範

### 整體風格
- **白底亮色主題**（類新聞媒體排版）
- body 背景：`#f0f4f8`
- 字型：`Noto Sans TC`
- 四欄等寬 grid，RWD 支援（≤1100px 改兩欄，≤640px 改單欄）

### 分類配色（CSS 變數）
每個分類欄用 `--cat-color`、`--cat-bg`、`--cat-border` 三個 CSS 變數控制：

| 分類 | `--cat-color` | `--cat-bg` | `--cat-border` |
|------|--------------|------------|----------------|
| 國際經濟 | `#b45309` | `rgba(180,83,9,0.06)` | `rgba(180,83,9,0.2)` |
| PC產業 | `#0369a1` | `rgba(3,105,161,0.06)` | `rgba(3,105,161,0.2)` |
| 上游廠商動態 | `#6d28d9` | `rgba(109,40,217,0.06)` | `rgba(109,40,217,0.2)` |
| 各品牌廠動態 | `#047857` | `rgba(4,120,87,0.06)` | `rgba(4,120,87,0.2)` |

### HTML 結構（每週 body 固定格式）

```html
<body>
  <!-- Header -->
  <div class="page-header">
    <div class="header-brand">
      <div class="header-logo">N</div>
      <div>
        <div class="header-title">ASUS News Intelligence</div>
        <div class="header-subtitle">{year}wk{XX} · PC 產業週報</div>
      </div>
    </div>

    <!-- 週次導覽 Pills（最新8週，最新在左，當週加 active） -->
    <div class="history-nav">
      <a href="index.html" class="week-pill active">2026wk27</a>
      <a href="2026wk26.html" class="week-pill">2026wk26</a>
      <!-- ... -->
    </div>

    <div class="header-date">{YYYY-MM-DD} 更新</div>
  </div>

  <!-- 四欄 Grid -->
  <div class="news-grid">

    <!-- 每一欄（套用 CSS 變數） -->
    <div class="news-col" style="--cat-color:#b45309; --cat-bg:rgba(180,83,9,0.06); --cat-border:rgba(180,83,9,0.2);">

      <!-- 欄標題 -->
      <div class="cat-header">
        <div class="cat-meta">
          <span class="cat-icon">📈</span>
          <span class="cat-label">國際經濟</span>
          <span class="cat-count">{新聞筆數}</span>
        </div>
        <div class="cat-summary">{本週該分類的 AI 總結標題}</div>
      </div>

      <!-- AI 觀點 -->
      <div class="ai-insight">
        <div class="ai-label">💡 AI 觀點</div>
        <div class="ai-text">{對華碩/筆電市場的影響分析}</div>
      </div>

      <!-- 新聞卡片（每篇一張） -->
      <div class="news-item">
        <div class="news-row">
          <h4 class="news-title">{新聞標題}</h4>
          <a href="{原文URL}" target="_blank" class="link-btn">🔗 原文</a>
        </div>
        <p class="news-body">{AI摘要，最多5行}</p>
        <span class="news-time">{發佈日期 YYYY-MM-DD}</span>
      </div>

      <!-- 更多 news-item... -->

    </div>
    <!-- 更多 news-col... -->

  </div>
</body>
```

### 週次 Pill 導覽規則
- 當週永遠 `href="index.html"`（不是 `2026wkXX.html`）
- 當週加 `class="week-pill active"`
- 歷史週連結為 `href="{year}wk{XX}.html"`
- 最多顯示最近 8 週
- 排序：最新週在最左

---

## 五、新聞過濾規則（`news_collector.py` 已實作）

以下類型會自動跳過，**更新時請維持這些規則**：

1. **購物/產品頁面網域**：`tw.buy.yahoo.com`、`shopping.pchome.com.tw`、`momoshop.com.tw`、`shopee.tw` 等
2. **URL 含產品路徑**：`/product/`、`/products/`、`/item/`、`/goods/` 等
3. **UUID 標題**：標題含 `xxxxxxxx-xxxx-xxxx` 格式的系統編碼
4. **產品規格標題**：標題含 `SSD`、`Win11`、`DDR`、`PCIe`、`GB` 等規格字串（疑似購物頁面）
5. **重複新聞**：已在 Excel 中的標題或連結

---

## 六、Excel 欄位說明

| 欄位 | 說明 |
|------|------|
| Year | 年份（如 2026） |
| Week | ISO 週次（如 27） |
| Category | 分類名稱（國際經濟/PC產業/上游廠商動態/各品牌廠動態） |
| Topic | AI 整理後的新聞標題 |
| Content | 新聞摘要 |
| Link | 原文連結 |
| 時間 | 發佈時間 |
| 完整內容 | 原文全文（若有抓取） |
| 對於筆電市場/華碩的影響 | AI 分析的影響評估 |

---

## 七、週次邏輯

- 從 Excel 讀取最大 Year，再取該 Year 的最大 Week
- 建議值 = max_week + 1（若超過 53 則進入下一年 week 1）
- **每次執行會互動確認**，可手動覆蓋
- 若 Excel 資料有誤（如出現 2027wkXX 錯誤資料），請先在 Excel 中手動刪除再執行

---

## 八、注意事項

1. 執行前確認 **Excel 已關閉**，否則 openpyxl 無法寫入
2. `generate_html_dashboard()` 使用 `try/except` 包住整個函數，若發生錯誤只會印警告、不會中斷程式，請注意 CMD 輸出有無 `[警訊]`
3. 歷史週報（如 `2026wk26.html`）一旦產生後不會自動更新，需要手動重新執行 `generate_html_only.py` 並重命名/覆蓋
4. 週次導覽 pills 的清單是從 `D:\ASUS\Anti-NotebookLM\NEWS\` 目錄掃描 `*wk*.html` 檔名自動生成的
