import os
import time
import pandas as pd
from news_collector import analyze_news_content, EXCEL_PATH, generate_html_dashboard

def main():
    print("============================================================")
    print(" 啟動 Excel 現有新聞「重新精簡化與 AI 分析」工具")
    print("============================================================")
    
    if not os.path.exists(EXCEL_PATH):
        print(f"[錯誤] 找不到 Excel 檔案：{EXCEL_PATH}")
        return

    print("-> 讀取 Excel 中...")
    df = pd.read_excel(EXCEL_PATH, sheet_name="News")
    
    # 找到最新的一週
    valid_rows = df[df["Topic"].astype(str).str.strip() != ""]
    if valid_rows.empty:
        print("沒有找到有效的新聞資料。")
        return
        
    first_row = valid_rows.iloc[0]
    year_val = int(float(first_row.get("Year", 0)))
    week_val = int(float(first_row.get("Week", 0)))
    
    print(f"-> 偵測到最新週數為 {year_val} 年 第 {week_val} 週")
    
    # 尋找屬於該週的列索引
    target_indices = []
    for idx, row in df.iterrows():
        try:
            if int(float(row.get("Year", 0))) == year_val and int(float(row.get("Week", 0))) == week_val:
                target_indices.append(idx)
        except:
            pass

    print(f"-> 共有 {len(target_indices)} 篇新聞需要重新進行極簡化 AI 分析...\n")
    
    for idx in target_indices:
        title = df.at[idx, "Topic"]
        full_content = df.at[idx, "完整內容"]
        
        # 排除完全無法取得內文的情況
        if not full_content or str(full_content).strip() == "" or "由於網站保護機制" in str(full_content) or str(full_content) == "nan":
            print(f"  [跳過] 缺乏完整內容：{title}")
            continue
            
        print(f"  [AI 處理中] 正在重新濃縮：{title} ...")
        
        analysis = analyze_news_content(str(full_content))
        if analysis:
            df.at[idx, "Topic"] = analysis.title
            df.at[idx, "Content"] = analysis.summary
            df.at[idx, "對於筆電市場/華碩的影響"] = analysis.impact_analysis
            print("   -> 成功精簡！")
        else:
            print("   -> AI 分析失敗，保留原內容。")
            
        time.sleep(1.5)  # 避免 API 限制
        
    print("\n-> 正在將更新後的極簡摘要寫回 Excel...")
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="News", index=False)
        print("[成功] Excel 更新完畢！")
    except PermissionError:
        print("\n[錯誤] 無法寫入！請先關閉您的 Excel 檔案後，重新執行此程式。")
        return

    # 重新生成 HTML
    local_html_path = r"D:\ASUS\Anti-NotebookLM\NEWS\index.html"
    print("\n-> 正在根據更新後的極簡資料，產生 HTML 看板...")
    generate_html_dashboard(EXCEL_PATH, local_html_path)
    
    import webbrowser
    if os.path.exists(local_html_path):
        webbrowser.open(local_html_path)
        print("[成功] 已在瀏覽器中自動為您開啟更新後的「ASUS 新聞情報看板」網頁！")

if __name__ == "__main__":
    main()
