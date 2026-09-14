import openpyxl

def main():
    file_path = r"D:\ASUS\News & Report\News\News.xlsx"
    print(f"正在開啟 Excel 檔案: {file_path}")
    
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    except Exception as e:
        print(f"開啟檔案失敗: {e}")
        print("請確認 Excel 檔案沒有被其他程式（如 Excel 軟體）開啟。")
        return

    insights = {
        "PC產業": "[AI觀點] 受到第三季售價調漲與品牌廠提前備貨影響，全球NB出貨動能短期承壓，且高單價的零組件持續墊高終端售價。對華碩而言，下半年市場需求可能面臨下滑挑戰，需謹慎控管庫存水位，並強化高階機種的銷售策略以維持整體獲利。",
        "上游廠商動態": "[AI觀點] 英特爾CPU再度漲價並縮減低階產線，加上記憶體大廠積極往先進製程與AI應用推進，進一步推升PC硬體成本。華碩應密切關注上游定價策略對利潤的擠壓效應，並可評估擴大與ARM架構晶片廠的合作，以豐富產品線並分散成本風險。",
        "各品牌廠動態": "[AI觀點] 聯想攜手Google推出深度整合Gemini模型的高階AI筆電，顯示AI PC的戰場已從硬體規格延伸至作業系統與軟體生態系的結合。華碩應持續深化自身AI應用的開發與跨平台整合，以在激烈的高階AI PC市場中鞏固競爭優勢。",
        "國際經濟": "[AI觀點] 國際原物料價格飆漲加劇通膨隱憂，同時新台幣展現強勁升值態勢，為外銷導向的科技業帶來匯兌壓力。華碩需靈活運用金融避險工具因應匯率波動，並審慎評估通膨對全球消費性電子終端買氣的潛在衝擊。"
    }

    # 尋找對應的欄位索引 (1-based)
    headers = {cell.value: i for i, cell in enumerate(ws[1], 1)}
    year_col = headers.get('Year')
    week_col = headers.get('Week')
    cat_col = headers.get('Category')
    impact_col = headers.get('對於筆電市場/華碩的影響')

    if not all([year_col, week_col, cat_col, impact_col]):
        print("找不到對應的欄位，請確認 Excel 第一列是否有 'Year', 'Week', 'Category', '對於筆電市場/華碩的影響' 標題。")
        return

    updated_count = 0
    for row in range(2, ws.max_row + 1):
        year = ws.cell(row=row, column=year_col).value
        week = ws.cell(row=row, column=week_col).value
        category = ws.cell(row=row, column=cat_col).value
        
        # 轉換為數字以防格式差異
        try:
            year = int(year) if year is not None else None
            week = int(week) if week is not None else None
        except ValueError:
            pass

        if year == 2026 and week == 38 and category in insights:
            ws.cell(row=row, column=impact_col).value = insights[category]
            updated_count += 1

    print(f"成功找到並填入了 {updated_count} 筆 wk38 的 AI 觀點。")
    
    try:
        print("正在儲存檔案...")
        wb.save(file_path)
        print("✅ Excel 檔案更新成功！")
    except Exception as e:
        print(f"儲存檔案失敗: {e}")
        print("請確認 Excel 檔案沒有被其他程式開啟，再試一次。")

if __name__ == "__main__":
    main()
