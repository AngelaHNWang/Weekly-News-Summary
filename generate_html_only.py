import os
import webbrowser
import pandas as pd
from datetime import datetime
from news_collector import generate_html_dashboard, EXCEL_PATH, auto_deploy_to_github

def main():
    print("============================================================")
    print(" 啟動純 HTML 看板產生器 (不抓取新新聞，直接讀取 Excel)")
    print("============================================================")
    
    local_html_path = r"D:\ASUS\Anti-NotebookLM\NEWS\index.html"
    
    if not os.path.exists(EXCEL_PATH):
        print(f"[錯誤] 找不到 Excel 檔案：{EXCEL_PATH}")
        return

    # 1. 偵測 Excel 中的最新週次
    suggested_year = datetime.now().year
    suggested_week = datetime.now().isocalendar()[1]
    try:
        all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
        news_sheets = [s for s in all_sheets.values() if "Topic" in s.columns or "Link" in s.columns]
        df = pd.concat(news_sheets, ignore_index=True) if news_sheets else pd.DataFrame()
        if not df.empty:
            full_df = df.fillna("")
            valid_rows = full_df[full_df["Topic"].astype(str).str.strip() != ""]
            if not valid_rows.empty:
                max_yr = int(valid_rows["Year"].replace("", 0).astype(float).max())
                yr_rows = valid_rows[valid_rows["Year"].replace("", 0).astype(float).astype(int) == max_yr]
                max_wk = int(yr_rows["Week"].replace("", 0).astype(float).max())
                if max_yr > 0:
                    suggested_year = max_yr
                if max_wk > 0:
                    suggested_week = max_wk
    except Exception as e:
        print(f"  [警訊] 自動偵測 Excel 最新週次失敗: {e}")

    print(f"\n  [偵測] Excel 最新資料為 {suggested_year}wk{suggested_week:02d}")
    print("  (您可以輸入指定的週次例如 28，或是直接按 Enter 以最上方資料為準)")
    ans = input(f"  請輸入欲產出的週次 (直接 Enter 使用 {suggested_week}，或輸入數字覆寫)：").strip()
    
    target_year = suggested_year
    target_week = suggested_week
    if ans.isdigit():
        target_week = int(ans)
        
    print(f"-> 正在讀取 Excel 檔案，產出週次：{target_year}wk{target_week:02d}...")
    generate_html_dashboard(EXCEL_PATH, local_html_path, target_year=target_year, target_week=target_week)
    
    if os.path.exists(local_html_path):
        print("-> HTML 網頁生成完畢！")
        try:
            webbrowser.open(local_html_path)
            print("[成功] 已在瀏覽器中自動為您開啟「ASUS 新聞情報看板」網頁！")
        except Exception as e:
            print(f"  [警訊] 自動開啟網頁時出錯: {e}")
            
        # 自動推送到 GitHub Pages
        auto_deploy_to_github()
    else:
        print("[錯誤] HTML 檔案未成功生成。")

if __name__ == "__main__":
    main()
