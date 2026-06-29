import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

EXCEL_PATH = r'D:\ASUS\News & Report\News\News.xlsx'

# 為 OECD 那篇補上 AI 觀點
impact_text = "[AI觀點] OECD下修全球經濟成長預測，反映貿易戰與地緣風險升溫，將抑制企業IT支出意願，對PC/NB整體出貨量形成壓力。華碩需留意歐美市場需求放緩風險，並加速布局高附加價值的AI PC產品以維持成長動能。"

wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb['News']

# Excel 第2列 = openpyxl row 2，第9欄 (I欄) = 對於筆電市場/華碩的影響
ws.cell(row=2, column=9, value=impact_text)
wb.save(EXCEL_PATH)
print(f"已為第2列 (OECD) 補上 AI 觀點")
print(f"內容: {impact_text}")
