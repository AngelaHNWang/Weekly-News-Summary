import pandas as pd
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

EXCEL_PATH = r'D:\ASUS\News & Report\News\News.xlsx'

# 讀取 Excel，檢查前12筆是否都有 AI 觀點
df = pd.read_excel(EXCEL_PATH, sheet_name='News')
wk24 = df.iloc[0:12]

missing = []
for i, row in wk24.iterrows():
    impact = str(row.get('對於筆電市場/華碩的影響', '')).strip()
    if impact == '' or impact == 'nan' or impact == 'NaN':
        missing.append(i)
        print(f"[缺少AI觀點] 第{i+2}列: {row.get('Topic','')}")

if not missing:
    print("全部12篇都已有 AI 觀點！")
else:
    print(f"\n共有 {len(missing)} 篇缺少 AI 觀點")
