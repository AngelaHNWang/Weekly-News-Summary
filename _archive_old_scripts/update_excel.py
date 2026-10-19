import pandas as pd
import os

excel_path = r'D:\ASUS\News & Report\News\News.xlsx'
df = pd.read_excel(excel_path)

# Find the row where Topic contains "OECD" and Category is "國際經濟"
mask = (df['Category'] == '國際經濟') & (df['Topic'].str.contains('OECD', na=False))
df.loc[mask, 'Topic'] = 'OECD下修今年全球成長至2.8%，能源供應成關鍵變數'

# Save back
df.to_excel(excel_path, index=False)
print("Updated Excel file successfully.")
