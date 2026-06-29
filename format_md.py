import pandas as pd

df = pd.read_excel(r'D:\ASUS\News & Report\News\News.xlsx')
# BUG FIXED: Added Year filter!
df_27 = df[(df['Week'] == 27) & (df['Year'] == 2026)].copy()

def map_cat(c):
    c = str(c)
    if '經' in c or 'Macro' in c: return '國際經濟 (Macro)'
    if 'PC' in c or '產業' in c: return 'PC產業 (Industry)'
    if '上游' in c or '供' in c: return '上游廠商動態 (Upstream / Supply Chain)'
    if '下游' in c or '品牌' in c: return '品牌廠動態 (Brand)'
    return c

df_27['Mapped_Cat'] = df_27['Category'].apply(map_cat)

with open(r'D:\ASUS\News & Report\News\2026Wk27_Formatted.md', 'w', encoding='utf-8') as f:
    f.write('# 2026Wk27 News Summary\n\n')
    f.write('這是一份為 2026wk27 準備的週報。請確保資訊圖表的標題完全一致，並分為四個區域。\n\n')
    
    order = ['國際經濟 (Macro)', 'PC產業 (Industry)', '品牌廠動態 (Brand)', '上游廠商動態 (Upstream / Supply Chain)']
    for target_cat in order:
        f.write(f'## {target_cat}\n')
        subset = df_27[df_27['Mapped_Cat'] == target_cat]
        if len(subset) == 0:
            f.write("本週無重大動態。\n\n")
        for _, row in subset.iterrows():
            f.write(f"### {row['Topic']}\n{row['Content']}\n\n")
