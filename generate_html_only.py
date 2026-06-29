import os
import webbrowser
from news_collector import generate_html_dashboard, EXCEL_PATH, auto_deploy_to_github

def main():
    print("============================================================")
    print(" 啟動純 HTML 看板產生器 (不抓取新新聞，直接讀取 Excel)")
    print("============================================================")
    
    local_html_path = r"D:\ASUS\Anti-NotebookLM\NEWS\index.html"
    
    if not os.path.exists(EXCEL_PATH):
        print(f"[錯誤] 找不到 Excel 檔案：{EXCEL_PATH}")
        return

    print(f"-> 正在讀取 Excel 檔案：{EXCEL_PATH}")
    generate_html_dashboard(EXCEL_PATH, local_html_path)
    
    if os.path.exists(local_html_path):
        print("-> HTML 網頁生成完畢！")
        try:
            webbrowser.open(local_index_path if 'local_index_path' in locals() else local_html_path)
            print("[成功] 已在瀏覽器中自動為您開啟「ASUS 新聞情報看板」網頁！")
        except Exception as e:
            print(f"  [警訊] 自動開啟網頁時出錯: {e}")
            
        # 自動推送到 GitHub Pages
        auto_deploy_to_github()
    else:
        print("[錯誤] HTML 檔案未成功生成。")

if __name__ == "__main__":
    main()
