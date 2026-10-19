import os
import html
import time
import webbrowser
import pandas as pd
from google import genai
from pydantic import BaseModel
from dotenv import load_dotenv

# ============================================================
#  ASUS 新聞情報 Infographic 生成器
#  直接從 Excel 讀取當週新聞資料，自動生成資訊圖表 HTML
# ============================================================

load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

# Excel 路徑
EXCEL_PATH = r"D:\ASUS\News & Report\News\News.xlsx"
# 輸出路徑
OUTPUT_PATH = r"D:\ASUS\Anti-NotebookLM\NEWS\infographic.html"

# 分類對應到區塊的映射
CATEGORY_TO_BLOCK = {
    "國際經濟": "A",
    "PC產業": "B",
    "上游廠商動態": "C",
    "各品牌廠動態": "D",
}

BLOCK_NAMES = {
    "A": "國際經濟 (Macro)",
    "B": "PC 產業 (Industry)",
    "C": "上游廠商動態 (Upstream / Supply Chain)",
    "D": "品牌廠動態 (Brand)",
}


# === AI 主題生成 ===
class BlockTheme(BaseModel):
    theme_a: str
    theme_b: str
    theme_c: str
    theme_d: str


def generate_themes(blocks_data):
    """根據四個區塊的新聞摘要，為每個區塊生成一句精練的重點標題。"""
    text_for_ai = ""
    for key in ["A", "B", "C", "D"]:
        items = blocks_data.get(key, [])
        text_for_ai += f"\n【區塊 {key}｜{BLOCK_NAMES[key]}】\n"
        for item in items:
            text_for_ai += f"- {item['summary']}\n"

    prompt = f"""
    請根據以下四個區塊的新聞摘要內容，為每一個區塊生成一句「重點標題 (Theme)」。
    要求：
    - 標題必須極簡，約 8~15 字以內
    - 精準涵蓋該區塊所有新聞的核心概念
    - 使用繁體中文
    
    內容：
    {text_for_ai}
    """
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config={
                    "response_mime_type": "application/json",
                    "response_schema": BlockTheme,
                },
            )
            return response.parsed
        except Exception as e:
            if "503" in str(e) and attempt < max_retries - 1:
                print(f"  [警訊] 生成主題時伺服器擁擠 (503)，等待 5 秒後進行第 {attempt + 2} 次重試...")
                time.sleep(5)
                continue
            print(f"  [錯誤] Theme generation failed: {e}")
            return None


def main():
    print("=" * 60)
    print(" ASUS 新聞情報 Infographic 生成器")
    print("=" * 60)

    # === 第一步：讀取 Excel 並篩選當週資料 ===
    print(f"\n-> 正在讀取 Excel 檔案：{EXCEL_PATH}")
    df = pd.read_excel(EXCEL_PATH, sheet_name="News").head(60).fillna("")

    # 找出最新一週的年份與週數
    valid_rows = df[df["Topic"].astype(str).str.strip() != ""]
    if valid_rows.empty:
        print("[錯誤] Excel 中沒有找到任何有效新聞資料！")
        return

    year_val = int(float(valid_rows.iloc[0].get("Year", 2026)))
    week_val = int(float(valid_rows.iloc[0].get("Week", 1)))
    print(f"   偵測到最新週次：{year_val} 年第 {week_val} 週")

    def is_current_week(r):
        try:
            return int(float(r.get("Year", 0))) == year_val and int(float(r.get("Week", 0))) == week_val
        except (ValueError, TypeError):
            return False

    df_current = df[df.apply(is_current_week, axis=1)]
    print(f"   找到 {len(df_current)} 篇當週新聞")

    # === 第二步：按分類整理到 A-D 區塊 ===
    blocks_data = {"A": [], "B": [], "C": [], "D": []}

    for _, row in df_current.iterrows():
        category = str(row.get("Category", "")).strip()
        block_key = CATEGORY_TO_BLOCK.get(category)
        if not block_key:
            print(f"  [警訊] 無法對應分類「{category}」到任何區塊，跳過此則")
            continue

        topic = str(row.get("Topic", "")).strip()
        summary = str(row.get("Content", "")).strip()
        impact = str(row.get("對於筆電市場/華碩的影響", "")).strip()
        link = str(row.get("Link", "")).strip()

        blocks_data[block_key].append({
            "topic": topic,
            "summary": summary,
            "impact": impact,
            "link": link,
        })

    for key in ["A", "B", "C", "D"]:
        print(f"   區塊 {key} ({BLOCK_NAMES[key]}): {len(blocks_data[key])} 篇")

    # === 第三步：AI 生成區塊主題標題 ===
    print("\n-> 正在使用 AI 生成各區塊主題標題...")
    themes = generate_themes(blocks_data)
    theme_dict = {}
    if themes:
        theme_dict["A"] = themes.theme_a
        theme_dict["B"] = themes.theme_b
        theme_dict["C"] = themes.theme_c
        theme_dict["D"] = themes.theme_d
        for key in ["A", "B", "C", "D"]:
            print(f"   [{key}] {theme_dict[key]}")
    else:
        theme_dict = {k: "主題生成失敗" for k in ["A", "B", "C", "D"]}

    # === 第四步：生成 HTML ===
    print("\n-> 正在生成視覺化 Infographic HTML...")
    header_title = f"2026Wk{week_val:02d} News Summary"

    html_blocks = ""
    for b_key in ["A", "B", "C", "D"]:
        items = blocks_data[b_key]

        html_blocks += f'''
        <div class="info-block">
            <div class="block-header block-{b_key}">
                <div class="block-badge">{b_key}</div>
                <div class="block-titles">
                    <h2 class="block-cat">{BLOCK_NAMES[b_key]}</h2>
                    <h3 class="block-theme">{html.escape(theme_dict[b_key])}</h3>
                </div>
            </div>
            <div class="block-content">
        '''

        if not items:
            html_blocks += '<p style="color: #94a3b8; padding: 1rem;">本週此分類暫無新聞。</p>'
        else:
            for item in items:
                # 處理 AI 觀點文字：移除 [AI觀點] 標籤、** 粗體標記
                impact_text = item["impact"]
                impact_text = impact_text.replace("[AI觀點]", "").replace("[AI觀點] ", "").strip()
                impact_text = impact_text.replace("**", "")

                link_btn = ""
                if item["link"] and item["link"].startswith("http"):
                    link_btn = f'<a href="{html.escape(item["link"])}" target="_blank" class="glass-btn">🔗 閱讀原文</a>'

                html_blocks += f'''
                <div class="info-card">
                    <div class="card-topic">
                        <span class="icon-topic">📰</span>
                        <h4>{html.escape(item["topic"])}</h4>
                    </div>
                    <div class="card-summary">
                        <span class="icon-summary">📝</span>
                        <p>{html.escape(item["summary"])}</p>
                    </div>
                    <div class="card-impact">
                        <span class="icon-impact">💡</span>
                        <p>{html.escape(impact_text)}</p>
                    </div>
                    <div class="card-link">
                        {link_btn}
                    </div>
                </div>
                '''

        html_blocks += '''
            </div>
        </div>
        '''

    # === 完整 HTML 模板 ===
    final_html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{header_title}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-dark: #07090f;
            --glass-bg: rgba(255, 255, 255, 0.03);
            --glass-border: rgba(255, 255, 255, 0.08);
            --primary: #00f2fe;
            --accent-a: #ff3366;
            --accent-b: #00f2fe;
            --accent-c: #a277ff;
            --accent-d: #ffaa00;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            background-color: var(--bg-dark);
            color: var(--text-main);
            font-family: 'Inter', 'Noto Sans TC', sans-serif;
            background-image: 
                radial-gradient(circle at 15% 50%, rgba(0, 242, 254, 0.08), transparent 25%),
                radial-gradient(circle at 85% 30%, rgba(255, 51, 102, 0.08), transparent 25%);
            background-attachment: fixed;
            min-height: 100vh;
            padding: 3rem 1rem;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        .main-title {{
            font-family: 'Outfit', sans-serif;
            font-size: 3.5rem;
            text-align: center;
            margin-bottom: 0.5rem;
            background: linear-gradient(135deg, #fff 0%, #a5b4fc 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-transform: uppercase;
            letter-spacing: 2px;
            text-shadow: 0 0 40px rgba(165, 180, 252, 0.2);
        }}
        .subtitle {{
            text-align: center;
            color: var(--text-muted);
            font-size: 1rem;
            margin-bottom: 3rem;
            letter-spacing: 1px;
        }}
        
        .info-block {{
            background: var(--glass-bg);
            border: 1px solid var(--glass-border);
            border-radius: 24px;
            padding: 2rem;
            margin-bottom: 3rem;
            backdrop-filter: blur(20px);
            box-shadow: 0 20px 40px rgba(0,0,0,0.4);
            transition: transform 0.3s ease, border-color 0.3s ease;
        }}
        .info-block:hover {{
            transform: translateY(-5px);
            border-color: rgba(255,255,255,0.15);
        }}
        
        .block-header {{
            display: flex;
            align-items: center;
            gap: 1.5rem;
            margin-bottom: 2rem;
            padding-bottom: 1.5rem;
            border-bottom: 1px solid var(--glass-border);
        }}
        .block-badge {{
            font-family: 'Outfit', sans-serif;
            font-size: 3rem;
            font-weight: 800;
            width: 80px;
            height: 80px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 20px;
            background: rgba(255,255,255,0.05);
            box-shadow: inset 0 0 20px rgba(255,255,255,0.05);
            flex-shrink: 0;
        }}
        .block-A .block-badge {{ color: var(--accent-a); text-shadow: 0 0 20px rgba(255,51,102,0.4); }}
        .block-B .block-badge {{ color: var(--accent-b); text-shadow: 0 0 20px rgba(0,242,254,0.4); }}
        .block-C .block-badge {{ color: var(--accent-c); text-shadow: 0 0 20px rgba(162,119,255,0.4); }}
        .block-D .block-badge {{ color: var(--accent-d); text-shadow: 0 0 20px rgba(255,170,0,0.4); }}
        
        .block-cat {{
            font-size: 1.1rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 0.5rem;
        }}
        .block-theme {{
            font-size: 1.8rem;
            font-weight: 600;
            color: #fff;
            line-height: 1.3;
        }}
        
        .block-content {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 1.5rem;
        }}
        
        .info-card {{
            background: rgba(0,0,0,0.2);
            border: 1px solid rgba(255,255,255,0.03);
            border-radius: 16px;
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            gap: 0.8rem;
            transition: background 0.3s, border-color 0.3s;
        }}
        .info-card:hover {{
            background: rgba(255,255,255,0.03);
            border-color: rgba(255,255,255,0.08);
        }}

        .card-topic {{
            display: flex;
            gap: 0.6rem;
            align-items: flex-start;
        }}
        .card-topic h4 {{
            font-size: 1rem;
            font-weight: 600;
            color: #e2e8f0;
            line-height: 1.4;
        }}
        
        .card-summary, .card-impact {{
            display: flex;
            gap: 0.8rem;
            align-items: flex-start;
        }}
        .card-summary p {{
            font-size: 0.9rem;
            line-height: 1.6;
            color: #cbd5e1;
            text-align: justify;
        }}
        .card-impact p {{
            font-size: 0.85rem;
            line-height: 1.5;
            color: #a78bfa;
            font-weight: 500;
            text-align: justify;
        }}
        .icon-topic, .icon-summary, .icon-impact {{
            font-size: 1.1rem;
            flex-shrink: 0;
            margin-top: 2px;
        }}
        
        .card-link {{
            margin-top: auto;
            padding-top: 0.8rem;
            text-align: right;
        }}
        .glass-btn {{
            display: inline-block;
            padding: 0.5rem 1.2rem;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 50px;
            color: #fff;
            text-decoration: none;
            font-size: 0.85rem;
            font-weight: 500;
            transition: all 0.2s;
        }}
        .glass-btn:hover {{
            background: rgba(255,255,255,0.15);
            transform: scale(1.05);
        }}
        
        @media (max-width: 768px) {{
            .main-title {{ font-size: 2.5rem; }}
            .block-header {{ flex-direction: column; text-align: center; gap: 1rem; }}
            .block-content {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1 class="main-title">{header_title}</h1>
        <p class="subtitle">ASUS 新聞情報看板 ── 本週 {len(df_current)} 則重點新聞</p>
        {html_blocks}
    </div>
</body>
</html>'''

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(final_html)

    print(f"\n[成功] Infographic HTML 已生成：{OUTPUT_PATH}")
    webbrowser.open(OUTPUT_PATH)
    print("[成功] 已在瀏覽器中自動開啟！")


if __name__ == "__main__":
    main()
