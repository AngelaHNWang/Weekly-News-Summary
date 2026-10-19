import os
import time
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv(r'D:\ASUS\Anti-NotebookLM\NEWS\.env')
API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=API_KEY)
client = genai

excel_path = r'D:\ASUS\News & Report\News\News.xlsx'
df = pd.read_excel(excel_path)

mask = (df['Year'] == 2026) & (df['Week'] == 41)
rows_to_update = df[mask]

print(f"Found {len(rows_to_update)} rows for 2026 wk41.")

for idx, row in rows_to_update.iterrows():
    if pd.isna(row.get('對於筆電市場/華碩的影響')) or str(row.get('對於筆電市場/華碩的影響')).strip() == "":
        content = row.get('完整內容')
        if pd.isna(content) or str(content).strip() == "":
            content = row.get('Content')
        
        if pd.isna(content) or str(content).strip() == "":
            print(f"Row {idx} has no content, skipping.")
            continue

        prompt = f"""你是一位專業的筆電產業分析師。
以下是一篇新聞內容：
{str(content)[:3000]}

請根據上述新聞內容，分析此事件對「筆記型電腦(PC/NB)市場」與「華碩(ASUS)」的影響。
要求：以「[AI觀點]」開頭，在 1-2 句話內、約 40-80 字精簡回答，使用繁體中文。"""
        
        try:
            print(f"Generating impact for: {row['Topic'][:30]}...")
            time.sleep(5)  # rate limit
            model = genai.GenerativeModel('gemini-flash-lite-latest')
            response = model.generate_content(prompt)
            impact = response.text.strip()
            df.at[idx, '對於筆電市場/華碩的影響'] = impact
            print(f" -> Success: {impact}")
        except Exception as e:
            print(f" -> Error: {e}")

# write back to excel
df.to_excel(excel_path, index=False)
print("Finished saving to Excel.")
