import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel(r'D:\ASUS\News & Report\News\News.xlsx', sheet_name='News')
wk24 = df.iloc[0:12].reset_index(drop=True)

# 先印出每篇的 Category 和 Topic 確認分布
print("=" * 60)
print(f"共 {len(wk24)} 筆資料，各分類篇數：")
print(wk24['Category'].value_counts().to_string())
print("=" * 60)

for i, row in wk24.iterrows():
    cat = row.get('Category','')
    topic = str(row.get('Topic',''))
    print(f"[{i+1:2d}] {cat:8s} | {topic[:80]}")

print("\n" + "=" * 60)
print("以下為各篇詳細內容：")
print("=" * 60)

for i, row in wk24.iterrows():
    print(f"\n--- [{i+1}] ---")
    print(f"Category: {row.get('Category','')}")
    print(f"Topic: {row.get('Topic','')}")
    content = str(row.get('Content',''))
    print(f"Content: {content[:400]}")
    print(f"Impact: {str(row.get('對於筆電市場/華碩的影響',''))[:250]}")
