import os
import sys
import time
import html

# 強制將標準輸出設為 utf-8，避免 Windows CMD 下的 cp950 編碼錯誤 (如簡體字或表情符號)
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        pass
import urllib.parse
import webbrowser
import random
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import xml.etree.ElementTree as ET

# 檢查依賴套件是否安裝，若無則友善提示
missing_packages = []
try:
    import pandas as pd
except ImportError:
    missing_packages.append("pandas")
try:
    import openpyxl
except ImportError:
    missing_packages.append("openpyxl")
try:
    import requests
except ImportError:
    missing_packages.append("requests")
try:
    from bs4 import BeautifulSoup
except ImportError:
    missing_packages.append("beautifulsoup4")
try:
    from google import genai
    from google.genai import types
except ImportError:
    missing_packages.append("google-genai")
try:
    from pydantic import BaseModel, Field
except ImportError:
    missing_packages.append("pydantic")
try:
    from googlenewsdecoder import gnewsdecoder
except ImportError:
    missing_packages.append("googlenewsdecoder")
try:
    from dotenv import load_dotenv
except ImportError:
    missing_packages.append("python-dotenv")

if missing_packages:
    print("=" * 60)
    print("[錯誤] 偵測到尚未安裝執行此工具所需的 Python 函式庫。")
    print("請開啟命令提示字元 (CMD) 並執行以下指令進行安裝：")
    print(f"pip install {' '.join(missing_packages)}")
    print("=" * 60)
    sys.exit(1)

# 載入環境變數
load_dotenv()
api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    print("=" * 60)
    print("[錯誤] 找不到 GEMINI_API_KEY 環境變數或 .env 檔案設定。")
    print("請至以下網址申請免費的 API Key:")
    print("https://aistudio.google.com/")
    print("並在工作區建立 .env 檔案填入：GEMINI_API_KEY=您的金鑰")
    print("=" * 60)
    sys.exit(1)

# 讀取 DIGITIMES 會員設定
DIGITIMES_USER = os.environ.get("DIGITIMES_USER")
DIGITIMES_PASSWORD = os.environ.get("DIGITIMES_PASSWORD")

# 初始化 Google GenAI 用戶端，使用最新的穩定版 gemini-2.5-flash
client = genai.Client(api_key=api_key)

# Excel 檔案路徑與設定
EXCEL_PATH = r"D:\ASUS\News & Report\News\News.xlsx"

# Google Alerts RSS 訂閱源設定 (方案一)
GOOGLE_ALERT_RSS_ECON = os.environ.get("GOOGLE_ALERT_RSS_ECON")


# 四大主題分類與對應 Google News 搜尋關鍵字 (加入排除投資、娛樂等字詞)
CATEGORIES = {
    "國際經濟": '("全球" OR "美國" OR "中國" OR "台灣" OR "日本" OR "英國" OR "法國" OR "德國" OR "義大利" OR "加拿大" OR "澳洲") 經濟 產業 OR 產品 -股票 -股價 -投資 -理財 -娛樂 -八卦 -影劇 -site:gov.tw -site:edu.tw -site:gov -site:edu',
    "PC產業": '("PC" OR "AIPC" OR "AI PC" OR "個人電腦" OR "筆記型電腦" OR "筆電") 市場 OR 產業 OR 產品 -股票 -股價 -投資 -理財 -娛樂 -八卦 -影劇',
    "上游廠商動態": '(Intel OR AMD OR Nvidia OR "記憶體" OR Micron OR 美光 OR SK海力士 OR 三星半導體 OR 晶圓代工 OR 台積電) 產業 OR 產品 -股票 -股價 -投資 -理財 -娛樂 -八卦 -影劇',
    "各品牌廠動態": '(Dell OR HP OR Lenovo OR Apple OR Acer OR MSI OR Samsung OR ASUS OR 華碩 OR 聯想 OR 戴爾 OR 惠普 OR 宏碁 OR 三星) PC OR 筆電 產業 OR 產品 -股票 -股價 -投資 -理財 -娛樂 -八卦 -影劇'
}

# 官方 RSS 來源 Adapter：針對特定分類額外訂閱高品質的官方新聞源 RSS。
# 相較於 Google News 搜尋結果，可取得更準確的原始發布時間、更少的側邊欄雜訊，且無需經過
# Google 轉址解密 (gnewsdecoder)，能降低額外的網路請求與被封鎖風險。
# 新增來源時只需在對應分類清單中加入 {"name":..., "url":..., "keywords":[...]}：
# keywords 用來從該來源的最新新聞中篩選出與本分類相關的標題，避免無關版面淹沒候選清單。
SOURCE_ADAPTERS = {
    "PC產業": [
        {"name": "iThome", "url": "https://www.ithome.com.tw/rss",
         "keywords": ["PC", "筆電", "筆記型電腦", "AI PC", "個人電腦", "NB"]},
    ],
    "上游廠商動態": [
        {"name": "TechNews", "url": "https://technews.tw/feed/",
         "keywords": ["Intel", "AMD", "Nvidia", "記憶體", "Micron", "美光", "SK海力士",
                       "三星半導體", "台積電", "晶圓", "半導體"]},
        {"name": "Tom's Hardware", "url": "https://www.tomshardware.com/feeds/all",
         "keywords": ["Intel", "AMD", "Nvidia", "memory", "chip", "foundry",
                       "Micron", "SK hynix", "TSMC", "semiconductor"]},
    ],
    "各品牌廠動態": [
        {"name": "iThome", "url": "https://www.ithome.com.tw/rss",
         "keywords": ["Dell", "HP", "Lenovo", "Apple", "Acer", "MSI", "Samsung",
                       "ASUS", "華碩", "聯想", "戴爾", "惠普", "宏碁", "三星"]},
    ],
}

# 每個類別預設抓取篇數
MAX_ARTICLES_PER_CATEGORY = 3

# Gemini API 呼叫併發上限：免費層級 RPM 依帳號/專案而異 (常見約 10-15 RPM，且 Google 已不再
# 公布統一數字，請至 https://aistudio.google.com/ 的專案額度頁確認實際值)。
# 這裡搭配各分類原有的 time.sleep(1) 保守設為同時最多 2 個併發請求，如果你的專案額度較高，
# 可以自行調高此數字以加快速度。
GEMINI_SEMAPHORE = threading.Semaphore(2)

# 無人值守模式：由 Windows 工作排程器等自動化程序以 `--auto` 參數啟動時開啟，
# 此模式下會跳過所有 input() 互動提示，避免排程執行時卡死等待輸入。
AUTO_MODE = "--auto" in sys.argv

# Pydantic 結構化輸出模型 (具備全文時使用)
class NewsItemAnalysis(BaseModel):
    title: str = Field(description="根據新聞內容提煉極簡短的主題標題(Topic)，讓讀者能馬上抓到重點，不要用原標題，限 5-15 字。")
    content_clean: str = Field(description="從網頁純文字中提取出的純淨新聞全文內文。應排除廣告、導覽列、側邊欄、版權宣告等雜訊")
    publish_time: str = Field(description="新聞的發布時間。若網頁中有提供日期時間請提取並統一轉換為 YYYY-MM-DD HH:MM 格式 (24小時制)；若只有日期沒有時間則輸出 YYYY-MM-DD；完全找不到日期時間時輸出空字串，禁止自行推測或臆測日期。")
    summary: str = Field(description="針對此篇新聞內容生成精簡的內文敘述，請精簡成 2-3 句話。")
    impact_analysis: str = Field(description="必須以『[AI觀點]』開頭 (不可省略、不可用其他文字取代)，評估對筆電市場/華碩的關鍵影響，控制在 1-2 句話、約 40-80 字，禁止使用 Markdown 語法 (如 ** 或 #)。")

# Pydantic 結構化輸出模型 (無內文、僅標題分析時使用)
class TitleOnlyAnalysis(BaseModel):
    title: str = Field(description="根據原新聞標題提煉極簡短的主題標題(Topic)，讓讀者能馬上抓到重點，不要用原標題，限 5-15 字。")
    summary: str = Field(description="根據標題生成背景說明，請精簡成 2-3 句話。")
    impact_analysis: str = Field(description="必須以『[AI觀點]』開頭 (不可省略、不可用其他文字取代)，評估對筆電市場/華碩的關鍵影響，控制在 1-2 句話、約 40-80 字，禁止使用 Markdown 語法 (如 ** 或 #)。")

# 防爬蟲：多組常見瀏覽器 User-Agent 輪替，降低被目標網站以固定 UA 識別並封鎖 (403) 的機率
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0",
]

def get_random_headers(referer="https://www.google.com/"):
    """
    隨機挑選一組瀏覽器 User-Agent 組成偽裝 headers，模擬使用者從搜尋引擎點擊進入，
    降低被目標網站以單一固定 UA 識別並封鎖的風險。
    """
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": referer,
    }

def polite_delay(min_seconds=0.5, max_seconds=2.0):
    """對外部網站發出請求前加入隨機微幅延遲，避免請求頻率過於規律而被判定為機器人"""
    time.sleep(random.uniform(min_seconds, max_seconds))

def get_digitimes_session(user, password):
    """
    使用 requests.Session() 模擬登入 DIGITIMES 台灣版網站以維持會員 Cookie
    """
    if not user or not password:
        return None
        
    print(f"  [資訊] 偵測到 DIGITIMES 會員設定，嘗試為帳號 {user[:4]}... 進行模擬登入...")
    session = requests.Session()
    # 同一個 session 生命週期內固定使用一組 UA，避免同一使用者中途切換瀏覽器指紋反而顯得可疑
    session.headers.update({
        "User-Agent": random.choice(USER_AGENTS),
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7"
    })

    login_url = "https://www.digitimes.com.tw/member/login.asp"
    try:
        # 先 GET 登入頁取得 cookie
        polite_delay()
        session.get(login_url, timeout=5)

        # 準備 POST payload
        payload = {
            "member_id": user,
            "member_pwd": password,
            "act": "login"
        }

        # 發送登入 POST
        polite_delay()
        response = session.post(login_url, data=payload, timeout=5)
        
        # 簡單驗證登入結果：如果內容中含有常見的登入失敗提示
        if "帳號不存在" in response.text or "密碼不正確" in response.text or "請輸入" in response.text:
            print("  [警訊] DIGITIMES 登入失敗：請確認 .env 中的帳密是否正確。將使用公開模式抓取。")
            return None
            
        print("  [成功] DIGITIMES 會員登入成功！已建立會員 Session。")
        return session
    except Exception as e:
        print(f"  [警訊] DIGITIMES 登入時發生異常，將使用公開模式抓取。錯誤: {e}")
        return None

def fetch_google_news_rss(query, max_results=3):
    """
    抓取 Google News RSS 並返回新聞的標題、臨時連結與發布時間
    """
    encoded_query = urllib.parse.quote(query)
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"

    try:
        polite_delay()
        response = requests.get(url, headers=get_random_headers(), timeout=5)
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        items = []
        for item in root.findall(".//item")[:max_results]:
            title = item.find("title").text
            link = item.find("link").text
            pub_date = item.find("pubDate").text
            items.append({
                "title": title,
                "link": link,
                "pub_date": pub_date
            })
        return items
    except Exception as e:
        print(f"  [警訊] 抓取 RSS 失敗 ({query}): {e}")
        return []

def fetch_google_alert_rss(rss_url, max_results=8):
    """
    抓取 Google 快訊 (Google Alerts) 的 RSS Feed，並返回新聞的標題、真實連結與發布時間。
    """
    try:
        polite_delay()
        response = requests.get(rss_url, headers=get_random_headers(), timeout=10)
        response.raise_for_status()

        # Google 快訊 RSS 是 Atom 格式
        root = ET.fromstring(response.content)
        
        # 定義 Atom namespace
        ns = {'atom': 'http://www.w3.org/2005/Atom'}
        entries = root.findall('atom:entry', ns)
        
        import re
        items = []
        for entry in entries[:max_results]:
            title_node = entry.find("atom:title", ns)
            link_node = entry.find("atom:link", ns)
            pub_node = entry.find("atom:published", ns)
            
            title_text = title_node.text if title_node is not None else ""
            # 清理標題中的 HTML 標籤（Google Alerts 常會把關鍵字加上 <b> 標籤）
            title_text = re.sub(r'<[^>]+>', '', title_text)
            
            raw_link = link_node.attrib.get("href") if link_node is not None else ""
            # 解析 Google Alerts 跳轉網址以取得真實網址
            real_link = raw_link
            if "google.com/url" in raw_link:
                try:
                    parsed = urllib.parse.urlparse(raw_link)
                    qs = urllib.parse.parse_qs(parsed.query)
                    if 'url' in qs:
                        real_link = qs['url'][0]
                except Exception as e:
                    print(f"  [警訊] 解析快訊跳轉網址失敗 ({raw_link}): {e}")
            
            pub_date = pub_node.text if pub_node is not None else ""
            
            items.append({
                "title": title_text.strip(),
                "link": real_link.strip(),
                "pub_date": pub_date
            })
        return items
    except Exception as e:
        print(f"  [警訊] 抓取 Google 快訊 RSS 失敗: {e}")
        return []

def fetch_news_with_fallback(base_query, max_results=3):
    """
    先嘗試當天 (24小時)，若新聞數量不足，則自動 backfill 過去 7 天內的新聞，確保本週前幾天的新聞不遺漏。
    """
    query_1d = f"{base_query} when:1d"
    items = fetch_google_news_rss(query_1d, max_results)
    
    if len(items) < max_results:
        print(f"  [資訊] 過去 24 小時內新聞不足，擴大搜尋範圍至過去 7 天以補齊...")
        query_7d = f"{base_query} when:7d"
        items_7d = fetch_google_news_rss(query_7d, max_results * 2)
        
        # 進行網址層級的合併去重
        seen_links = set(item["link"] for item in items)
        for item in items_7d:
            if item["link"] not in seen_links:
                items.append(item)
                seen_links.add(item["link"])
                
    return items[:max_results * 2]

def fetch_adapter_rss(source, max_results=15):
    """
    直接抓取指定的官方 RSS 來源 (SOURCE_ADAPTERS 中設定的 Adapter)，
    並僅保留標題中包含該分類關鍵字的候選新聞，避免來源網站的無關版面淹沒候選清單。
    由於連結已是真實發布者網址 (非 Google 轉址)，回傳的 item 會標記 "direct": True，
    讓主流程略過不必要的 decode_url 解碼步驟。
    """
    url = source["url"]
    keywords = source.get("keywords", [])
    name = source.get("name", url)
    try:
        polite_delay()
        response = requests.get(url, headers=get_random_headers(), timeout=8)
        response.raise_for_status()

        root = ET.fromstring(response.content)
        items = []
        for item in root.findall(".//item"):
            title_node = item.find("title")
            link_node = item.find("link")
            pub_node = item.find("pubDate")

            title_text = (title_node.text or "").strip() if title_node is not None else ""
            link_text = (link_node.text or "").strip() if link_node is not None else ""
            pub_text = (pub_node.text or "").strip() if pub_node is not None else ""

            if not title_text or not link_text:
                continue

            # 僅保留標題命中該分類關鍵字的新聞
            if keywords and not any(k.lower() in title_text.lower() for k in keywords):
                continue

            items.append({"title": title_text, "link": link_text, "pub_date": pub_text, "direct": True})
            if len(items) >= max_results:
                break
        return items
    except Exception as e:
        print(f"  [警訊] 抓取官方 RSS 來源失敗 ({name}): {e}")
        return []

def decode_url(google_url):
    """
    使用 googlenewsdecoder 庫解密 Google News 的跳轉 URL 以獲取原始發布者的真實連結
    """
    try:
        decoded_info = gnewsdecoder(google_url, interval=1)
        if decoded_info.get("status"):
            return decoded_info["decoded_url"]
    except Exception as e:
        print(f"  [警訊] URL 解碼失敗: {e}")
    return google_url

def clean_url(url):
    """
    移除網址中常見的追蹤用問號參數 (如 utm_source、fbclid 等)、錨點 (#)、尾部斜線與尾部隨機數字，
    以防去重機制失效。注意：只移除已知的追蹤參數，保留其餘查詢參數 (例如 DIGITIMES 用
    ?id=0000766 當作文章識別碼，並非追蹤噪音，若整段問號參數都砍掉會讓不同文章被誤判為重複)。
    """
    import re
    if not url:
        return ""

    _TRACKING_PARAMS = {
        "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "utm_id",
        "fbclid", "gclid", "ref", "ref_src", "spm", "from",
    }
    parsed = urllib.parse.urlparse(url)
    kept_qs = [(k, v) for k, v in urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
               if k.lower() not in _TRACKING_PARAMS]
    url = urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(kept_qs), fragment=""))

    # 移除尾部斜線
    url = url.rstrip("/")
    # 移除尾部的隨機數字 (例如 TechNews 網址結尾的 /12345)
    url = re.sub(r'/\d+$', '', url)
    return url.strip()


def extract_webpage_text(url, session=None):
    """
    抓取指定網頁，過濾無效的 CSS/JS，僅回傳網頁的純文字內容。
    如果是 DIGITIMES 的新聞且提供了登入的 session，將使用該 session 下載以獲取會員全文。
    """
    max_attempts = 2
    for attempt in range(1, max_attempts + 1):
        try:
            polite_delay()
            if "digitimes.com.tw" in url and session is not None:
                response = session.get(url, timeout=5)
            else:
                response = requests.get(url, headers=get_random_headers(), timeout=5, verify=True)

            if response.status_code == 403 and attempt < max_attempts:
                print(f"  [警訊] 遭遇 403 封鎖，更換 User-Agent 後重試一次 ({url[:60]}...)")
                polite_delay(1.5, 3.0)
                continue

            response.raise_for_status()

            if response.encoding == 'ISO-8859-1' or response.encoding is None:
                response.encoding = response.apparent_encoding

            soup = BeautifulSoup(response.text, 'html.parser')

            # 1. 移除無關的側邊欄、廣告推薦等區塊噪聲，只保留新聞主體
            noise_selectors = [
                '.sidebar', '#sidebar', '.aside', '.related-posts', '.popular-posts',
                '.entry-meta', '.social-share', '.comments-area', '.widget',
                '.trending', '.hot-news', '.header-menu', '.footer-container',
                '.post-ratings', '.recommend-posts', '.author-bio', '#comments',
                '.tagcloud', '.post-nav'
            ]
            for selector in noise_selectors:
                try:
                    for element in soup.select(selector):
                        element.extract()
                except Exception:
                    pass

            # 2. 移除標準 HTML 區塊標籤
            for element in soup(["script", "style", "nav", "footer", "header", "aside", "noscript"]):
                element.extract()

            text = soup.get_text(separator='\n')
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            cleaned_text = '\n'.join(lines)

            return cleaned_text[:20000]
        except Exception as e:
            if attempt < max_attempts:
                continue
            print(f"  [警訊] 抓取網頁內文失敗 ({url}): {e}")
            return ""
    return ""

def analyze_news_content(html_text):
    """
    使用 Gemini 2.5-flash 進行新聞的摘要、影響分析與結構化 JSON 輸出
    """
    prompt = f"""
請幫我分析以下這段網頁內容，它是一篇新聞報導。
網頁清理後的純文字如下：
---
{html_text}
---
請幫我提取並整理出：
1. 精簡主題標題 (title) - 根據新聞內容提煉出極簡短的主題標題(Topic)，讓讀者能馬上抓到重點，不要照抄冗長的原標題，字數限 5-15 字以內。
2. 去除雜訊後的新聞全文內文 (content_clean) - 提取真實新聞正文，排除網頁廣告、選單、版權宣告。
3. 新聞發布時間 (publish_time) - 從網頁內容中尋找實際刊登日期時間並統一轉換為 YYYY-MM-DD HH:MM 格式 (24小時制)；只有日期沒有時間則輸出 YYYY-MM-DD；完全找不到時輸出空字串，不要自行推測。
4. 內容摘要 (summary) - 內文敘述請精簡成 2-3 句話，直陳事實、數據與核心事件。
5. 華碩與筆電市場影響分析 (impact_analysis) - 評估該事件對「整個筆記型電腦 (PC/NB) 市場」與「華碩 (ASUS)」的關鍵影響，必須以「[AI觀點]」開頭，控制在 1-2 句話、約 40-80 字，不要使用 Markdown 語法。
   範例："[AI觀點] 記憶體價格上漲將壓縮筆電代工毛利，華碩可能需要在下一季調漲售價或選擇性犧牲部分入門機型的毛利以維持市佔。"

請嚴格遵守回傳的 JSON 格式 Schema，所有文字均須以繁體中文 (Traditional Chinese) 回答。
"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            with GEMINI_SEMAPHORE:
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt,
                    config={
                        'response_mime_type': 'application/json',
                        'response_schema': NewsItemAnalysis,
                        'temperature': 0.2,
                    }
                )
            return response.parsed
        except Exception as e:
            err_msg = str(e)
            if "503" in err_msg and attempt < max_retries - 1:
                print(f"  [警訊] Gemini 伺服器滿載 (503)，等待 5 秒後進行第 {attempt+2} 次重試...")
                time.sleep(5)
                continue
                
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                print("  [警訊] Gemini API 額度已達上限 (429 RESOURCE_EXHAUSTED)，無法進行 AI 分析。")
            elif "400" in err_msg or "API_KEY_INVALID" in err_msg:
                print("  [警訊] Gemini API 金鑰無效，請檢查 .env 設定。")
            else:
                print(f"  [警訊] Gemini 分析失敗: {err_msg[:100]}...")
            return None

def analyze_news_from_title(title, category):
    """
    當無法抓取網頁全文時，使用新聞標題結合 Gemini 2.5-flash 的知識庫進行背景與影響摘要
    """
    prompt = f"""
您是一位專業的科技與財經分析師。
目前我們有一篇新聞的標題為：「{title}」，屬於「{category}」主題。
由於目前網路抓取受到限制，請您光憑這個「新聞標題」，並結合您的背景知識，為我們撰寫：
1. 精簡主題標題 (title) - 根據原新聞標題提煉出極簡短的主題標題(Topic)，讓讀者能馬上抓到重點，不要照抄冗長的原標題，字數限 5-15 字以內。
2. 新聞重點背景說明 (summary) - 內文敘述請精簡成 2-3 句話，直陳事實與重點。
3. 華碩與筆電市場影響分析 (impact_analysis) - 評估本則新聞對「筆電市場 (PC/NB)」與「華碩 (ASUS)」的潛在影響，必須以「[AI觀點]」開頭，控制在 1-2 句話、約 40-80 字，不要使用 Markdown 語法。
   範例："[AI觀點] 記憶體價格上漲將壓縮筆電代工毛利，華碩可能需要在下一季調漲售價或選擇性犧牲部分入門機型的毛利以維持市佔。"

要求：
- 語氣需客觀、專業且肯定，不要使用「因為無法抓取內文」、「根據標題猜測」、「AI 預測」或「由於限制」等任何與系統限制相關的免責字句，直接產出分析內容即可。

請嚴格遵守回傳的 JSON 格式 Schema，所有文字均須以繁體中文 (Traditional Chinese) 回答。
"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            with GEMINI_SEMAPHORE:
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt,
                    config={
                        'response_mime_type': 'application/json',
                        'response_schema': TitleOnlyAnalysis,
                        'temperature': 0.2,
                    }
                )
            return response.parsed
        except Exception as e:
            err_msg = str(e)
            if "503" in err_msg and attempt < max_retries - 1:
                print(f"  [警訊] Gemini 伺服器滿載 (503)，等待 5 秒後進行第 {attempt+2} 次重試...")
                time.sleep(5)
                continue
                
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                print("  [警訊] Gemini API 額度已達上限 (429 RESOURCE_EXHAUSTED)，無法進行 AI 標題分析。")
            elif "400" in err_msg or "API_KEY_INVALID" in err_msg:
                print("  [警訊] Gemini API 金鑰無效，請檢查 .env 設定。")
            else:
                print(f"  [警訊] Gemini 標題分析失敗: {err_msg[:100]}...")
            return None

def generate_html_dashboard(excel_path, html_path, target_year=None, target_week=None):
    """
    從 Excel 讀取最新或指定週次的新聞，並生成具有美學動態效果的 HTML 互動看板。
    會過濾無效空白行，並動態將週數呈現在標題上。
    """
    try:
        if not os.path.exists(excel_path):
            return
            
        df = pd.read_excel(excel_path, sheet_name="News")
        if df.empty:
            return
            
        # 讀取最前 60 筆以確保在過濾空白行後能有足夠的資料量展示
        latest_df = df.head(60).fillna("")
        
        # 1. 決定動態週數標題：找到最新一筆有效的新聞列
        valid_rows = latest_df[latest_df["Topic"].astype(str).str.strip() != ""]
        if not valid_rows.empty:
            first_row = valid_rows.iloc[0]
            try:
                year_val = int(first_row.get("Year", datetime.now().year))
            except:
                year_val = datetime.now().year
            try:
                week_val = int(first_row.get("Week", datetime.now().isocalendar()[1]))
            except:
                week_val = datetime.now().isocalendar()[1]
        else:
            year_val = datetime.now().year
            week_val = datetime.now().isocalendar()[1]
            
        # 如果手動指定了週次或年份，則進行覆寫
        if target_year is not None:
            year_val = target_year
        if target_week is not None:
            week_val = target_week
            
        header_title = f"{year_val}wk{week_val:02d} News Summary"
        
        # 只保留與最新一筆週數相同的資料，避免混入上一週的新聞
        def is_current_week(r):
            try:
                return int(float(r.get("Year", 0))) == year_val and int(float(r.get("Week", 0))) == week_val
            except:
                return False
        latest_df = latest_df[latest_df.apply(is_current_week, axis=1)]
        
        categories = ["國際經濟", "PC產業", "上游廠商動態", "各品牌廠動態"]
        news_data = {cat: [] for cat in categories}
        
        # 2. 遍歷數據並進行有效性過濾
        for _, row in latest_df.iterrows():
            cat = str(row.get("Category", "")).strip()
            topic = str(row.get("Topic", "")).strip()
            
            # 過濾空白無效列 (若 Topic 為空則跳過)
            if not topic or topic.lower() == "nan":
                continue
                
            if cat in news_data:
                # 限制每個分類在前端只取最前 5 筆展示以求乾淨緊湊
                if len(news_data[cat]) >= 5:
                    continue
                    
                raw_time = str(row.get("時間", ""))
                clean_time = raw_time.split("\n")[0] if "\n" in raw_time else raw_time
                
                # === 使用者指定的標題客製化覆寫 ===
                title_text = str(topic)
                if "OECD" in title_text and cat == "國際經濟":
                    title_text = "OECD下修今年全球成長至2.8%，能源供應成關鍵變數"
                
                news_data[cat].append({
                    "title": html.escape(title_text),
                    "summary": html.escape(str(row.get("Content", ""))).replace("\n", "<br>"),
                    "impact": html.escape(str(row.get("對於筆電市場/華碩的影響", ""))).replace("\n", "<br>"),
                    "link": str(row.get("Link", "")),
                    "time": html.escape(clean_time),
                    "full_content": html.escape(str(row.get("完整內容", ""))).replace("\n", "<br>")
                })

        # 生成各類別的新聞條目 HTML
        cat_icons = {
            "國際經濟": "📈",
            "PC產業": "💻",
            "上游廠商動態": "⚙️",
            "各品牌廠動態": "🏷️"
        }
        
        # 使用 Gemini 為每個分類生成一句話標題
        def generate_section_headline(cat_name, items):
            """呼叫 Gemini 根據該分類下所有新聞標題，生成一句話總結標題"""
            if not items:
                return cat_name
            titles = "、".join([item["title"] for item in items])
            prompt = f"""你是一位專業的科技產業分析師。
以下是「{cat_name}」分類下本週的所有新聞標題：
{titles}

請用一句話（20-35字以內）歸納這些新聞的核心主題，作為本週此分類的總標題。
要求：直接輸出標題文字，不要加引號、不要加標點符號結尾、不要解釋。"""
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt
                    )
                    headline = response.text.strip().strip('"').strip('「').strip('」').strip('。').strip()
                    return headline if headline else cat_name
                except Exception as e:
                    if "503" in str(e) and attempt < max_retries - 1:
                        print(f"  [警訊] 生成 {cat_name} 標題時伺服器擁擠 (503)，等待 5 秒後進行第 {attempt+2} 次重試...")
                        time.sleep(5)
                        continue
                    print(f"  [警訊] 生成 {cat_name} 標題失敗: {e}")
                    return cat_name
        
        # 使用 Gemini 合併所有 AI 觀點為一段總結
        def generate_consolidated_impact(cat_name, items):
            """呼叫 Gemini 將該分類下所有新聞的 AI 觀點合併為一段總結"""
            if not items:
                return ""
            impacts = [item["impact"] for item in items if item["impact"] and item["impact"].strip() and item["impact"].strip().lower() != "nan"]
            if not impacts:
                return ""
            all_impacts = "\n".join(impacts)
            prompt = f"""你是一位專業的筆電產業分析師。
以下是「{cat_name}」分類下本週所有新聞的個別 AI 觀點（對筆電市場/華碩的影響）：
---
{all_impacts}
---

請將上述所有觀點整合為一段極度精簡的總結分析（1-3句話以內，用字精煉），說明這些新聞綜合來看對「筆記型電腦市場」與「華碩(ASUS)」的關鍵影響。
要求：以「[AI觀點]」開頭，語氣客觀專業且科技感，使用繁體中文。"""
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt
                    )
                    return response.text.strip() if response.text else ""
                except Exception as e:
                    if "503" in str(e) and attempt < max_retries - 1:
                        print(f"  [警訊] 合併 {cat_name} AI觀點時伺服器擁擠 (503)，等待 5 秒後進行第 {attempt+2} 次重試...")
                        time.sleep(5)
                        continue
                    print(f"  [警訊] 合併 {cat_name} AI觀點失敗: {e}")
                    # fallback: 如果失敗，只取前兩個觀點，移除重複的標籤並以條列呈現
                    cleaned = []
                    for imp in impacts[:2]:
                        imp = imp.replace("[AI觀點]", "").strip()
                        cleaned.append(f"<li style='margin-bottom: 0.3rem;'>{imp}</li>")
                    return "[AI觀點] <ul style='padding-left: 1.2rem; margin-top: 0.4rem; margin-bottom: 0;'>" + "".join(cleaned) + "</ul>"
        
        # 為每個分類生成標題與總結 AI 觀點
        print("-> 正在為各分類生成標題與 AI 觀點總結...")
        section_headlines = {}
        section_impacts = {}
        for cat in categories:
            items = news_data[cat]
            if items:
                section_headlines[cat] = generate_section_headline(cat, items)
                print(f"  [{cat}] 標題: {section_headlines[cat]}")
                section_impacts[cat] = generate_consolidated_impact(cat, items)
                time.sleep(1)
            else:
                section_headlines[cat] = cat
                section_impacts[cat] = ""
        
        # 4欄版面：將所有分類並排（各分類獨立配色）
        def build_four_columns_html():
            cat_colors = {
                "國際經濟":    {"main": "#b45309", "bg": "rgba(180,83,9,0.06)",   "border": "rgba(180,83,9,0.2)"},
                "PC產業":     {"main": "#0369a1", "bg": "rgba(3,105,161,0.06)",   "border": "rgba(3,105,161,0.2)"},
                "上游廠商動態": {"main": "#6d28d9", "bg": "rgba(109,40,217,0.06)", "border": "rgba(109,40,217,0.2)"},
                "各品牌廠動態": {"main": "#047857", "bg": "rgba(4,120,87,0.06)",   "border": "rgba(4,120,87,0.2)"},
            }
            col_html = '<div class="news-grid">'
            for cat in categories:
                items = news_data[cat]
                icon = cat_icons.get(cat, '📰')
                headline = section_headlines.get(cat, cat)
                impact_summary = section_impacts.get(cat, "")
                c = cat_colors.get(cat, cat_colors["PC產業"])
                main_color = c["main"]
                bg_color = c["bg"]
                border_color = c["border"]

                col_html += f'<div class="news-col" style="--cat-color:{main_color}; --cat-bg:{bg_color}; --cat-border:{border_color};">'

                # ── 欄標題區
                col_html += f'''
                <div class="cat-header">
                    <div class="cat-meta">
                        <span class="cat-icon">{icon}</span>
                        <span class="cat-label">{cat}</span>
                        <span class="cat-count">{len(items)}</span>
                    </div>
                    <div class="cat-summary">{headline}</div>
                </div>
                '''

                # ── AI 觀點區塊
                if impact_summary:
                    impact_clean = impact_summary.replace("[AI觀點]", "").replace("**", "").strip()
                    col_html += f'''
                    <div class="ai-insight">
                        <div class="ai-label">💡 AI 觀點</div>
                        <div class="ai-text">{impact_clean}</div>
                    </div>
                    '''

                # ── 新聞列表
                if not items:
                    col_html += '<p class="no-news">此分類本週尚無新聞。</p>'
                else:
                    for item in items:
                        link_btn = ""
                        if item["link"] and item["link"].startswith("http"):
                            link_btn = f'<a href="{item["link"]}" target="_blank" class="link-btn">🔗 原文</a>'
                        time_chip = ""
                        t = item.get("time", "")
                        if t and t != "nan" and t.strip():
                            time_chip = f'<span class="news-time">{t[:10]}</span>'
                        col_html += f'''
                        <div class="news-item">
                            <div class="news-row">
                                <h4 class="news-title">{item['title']}</h4>
                                {link_btn}
                            </div>
                            <p class="news-body">{item['summary']}</p>
                            {time_chip}
                        </div>
                        '''

                col_html += '</div>'
            col_html += '</div>'
            return col_html
            
        grid_html = build_four_columns_html()
        history_html = ""
        try:
            import glob
            history_files = glob.glob(os.path.join(os.path.dirname(html_path), "*wk*.html"))
            weeks = []
            for f in history_files:
                basename = os.path.basename(f)
                if basename.endswith(".html") and "wk" in basename:
                    w = basename.replace(".html", "")
                    if w not in weeks:
                        weeks.append(w)
            
            # 確保當前週數也在清單中
            current_wk_str = f"{year_val}wk{week_val:02d}"
            if current_wk_str not in weeks:
                weeks.append(current_wk_str)
                
            # 排序（新的在最前面）
            weeks.sort(reverse=True)
            
            if weeks:
                pills = []
                for w in weeks[:8]:  # 最多顯示 8 週
                    active_cls = "week-pill active" if w == current_wk_str else "week-pill"
                    val = f"{w}.html"
                    pills.append(f'<a href="{val}" class="{active_cls}">{w}</a>')
                history_html = f'<div class="history-nav">{"".join(pills)}</div>'
        except Exception as e:
            print(f"產生歷史選單錯誤: {e}")

        # 完整的 HTML 看板模板 (NotebookLM Infographic 風格)
        html_template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{header_title}</title>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        *, *::before, *::after {{ margin: 0; padding: 0; box-sizing: border-box; }}

        body {{
            background: #f0f4f8;
            color: #1e293b;
            font-family: 'Noto Sans TC', system-ui, sans-serif;
            line-height: 1.5;
            min-height: 100vh;
        }}

        /* ── Header ── */
        .page-header {{
            background: #ffffff;
            border-bottom: 1px solid #e2e8f0;
            padding: 0.75rem 1.5rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 1px 4px rgba(0,0,0,0.07);
        }}
        .header-brand {{
            display: flex;
            align-items: center;
            gap: 0.6rem;
            flex-shrink: 0;
        }}
        .header-logo {{
            width: 28px; height: 28px;
            background: #1e40af;
            border-radius: 6px;
            display: flex; align-items: center; justify-content: center;
            color: #fff; font-size: 0.85rem; font-weight: 800;
        }}
        .header-title {{
            font-size: 1rem;
            font-weight: 700;
            color: #1e293b;
        }}
        .header-subtitle {{
            font-size: 0.75rem;
            color: #64748b;
        }}
        .history-nav {{
            display: flex;
            align-items: center;
            gap: 0.35rem;
            flex-wrap: wrap;
        }}
        .week-pill {{
            display: inline-block;
            padding: 0.25rem 0.65rem;
            border-radius: 20px;
            font-size: 0.72rem;
            font-weight: 600;
            text-decoration: none;
            border: 1px solid #cbd5e1;
            color: #475569;
            background: #f8fafc;
            transition: all 0.15s;
            white-space: nowrap;
        }}
        .week-pill:hover {{ border-color: #1e40af; color: #1e40af; background: #eff6ff; }}
        .week-pill.active {{ background: #1e40af; color: #fff; border-color: #1e40af; }}
        .header-date {{ font-size: 0.72rem; color: #94a3b8; flex-shrink: 0; }}

        /* ── Grid ── */
        .news-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 1rem;
            padding: 1rem 1.25rem 1.5rem;
            max-width: 1800px;
            margin: 0 auto;
        }}
        .news-col {{
            display: flex;
            flex-direction: column;
            gap: 0.75rem;
            min-width: 0;
        }}

        /* ── Cat header ── */
        .cat-header {{
            background: #fff;
            border-radius: 10px;
            padding: 0.8rem 1rem;
            border: 1px solid #e2e8f0;
            border-top: 3px solid var(--cat-color);
        }}
        .cat-meta {{ display: flex; align-items: center; gap: 0.45rem; margin-bottom: 0.4rem; }}
        .cat-icon {{ font-size: 1rem; line-height: 1; }}
        .cat-label {{ font-size: 0.7rem; font-weight: 800; letter-spacing: 1px; color: var(--cat-color); }}
        .cat-count {{
            margin-left: auto;
            background: var(--cat-color); color: #fff;
            font-size: 0.65rem; font-weight: 700;
            padding: 0.1rem 0.45rem; border-radius: 12px;
        }}
        .cat-summary {{ font-size: 0.82rem; font-weight: 600; color: #334155; line-height: 1.45; }}

        /* ── AI insight ── */
        .ai-insight {{
            background: var(--cat-bg);
            border: 1px solid var(--cat-border);
            border-left: 3px solid var(--cat-color);
            border-radius: 8px;
            padding: 0.7rem 0.85rem;
        }}
        .ai-label {{ font-size: 0.65rem; font-weight: 800; color: var(--cat-color); letter-spacing: 0.8px; margin-bottom: 0.3rem; }}
        .ai-text {{ font-size: 0.8rem; color: #374151; line-height: 1.6; }}

        /* ── News items ── */
        .news-item {{
            background: #fff;
            border: 1px solid #e2e8f0;
            border-left: 3px solid var(--cat-color);
            border-radius: 8px;
            padding: 0.7rem 0.85rem;
            transition: box-shadow 0.15s, transform 0.15s;
        }}
        .news-item:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.09); transform: translateY(-1px); }}
        .news-row {{
            display: flex; align-items: flex-start;
            justify-content: space-between; gap: 0.4rem; margin-bottom: 0.35rem;
        }}
        .news-title {{ font-size: 0.88rem; font-weight: 700; color: #1e293b; line-height: 1.4; flex: 1; }}
        .link-btn {{
            flex-shrink: 0;
            background: var(--cat-bg); color: var(--cat-color);
            border: 1px solid var(--cat-border);
            padding: 0.15rem 0.45rem; border-radius: 4px;
            font-size: 0.63rem; font-weight: 600;
            text-decoration: none; white-space: nowrap; transition: all 0.15s;
        }}
        .link-btn:hover {{ background: var(--cat-color); color: #fff; }}
        .news-body {{
            font-size: 0.78rem; color: #4b5563; line-height: 1.65;
            display: -webkit-box; -webkit-line-clamp: 5;
            -webkit-box-orient: vertical; overflow: hidden;
        }}
        .news-time {{
            display: inline-block; margin-top: 0.4rem;
            font-size: 0.62rem; color: #94a3b8;
            background: #f1f5f9; padding: 0.1rem 0.4rem; border-radius: 3px;
        }}
        .no-news {{ font-size: 0.8rem; color: #94a3b8; padding: 0.5rem; }}

        /* ── Responsive ── */
        @media (max-width: 1100px) {{ .news-grid {{ grid-template-columns: repeat(2, 1fr); }} }}
        @media (max-width: 640px) {{
            .news-grid {{ grid-template-columns: 1fr; padding: 0.75rem; }}
            .page-header {{ flex-wrap: wrap; }}
        }}
        @media print {{ body {{ background: #fff; }} .page-header {{ box-shadow: none; position: static; }} }}
    </style>
</head>
<body>
    <div class="page-header">
        <div class="header-brand">
            <div class="header-logo">N</div>
            <div>
                <div class="header-title">ASUS News Intelligence</div>
                <div class="header-subtitle">{year_val}wk{week_val:02d} · PC 產業週報</div>
            </div>
        </div>
        {history_html}
        <div class="header-date">{datetime.now().strftime('%Y-%m-%d')} 更新</div>
    </div>
    {grid_html}
</body>
</html>
"""
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_template)
            
        current_wk_str = f"{year_val}wk{week_val:02d}"
        archive_path = os.path.join(os.path.dirname(html_path), f"{current_wk_str}.html")
        with open(archive_path, "w", encoding="utf-8") as f:
            f.write(html_template)
            
        print(f"成功更新 HTML 看板網頁：{html_path}")
        print(f"成功儲存歷史歸檔：{archive_path}")
        
        # 4. 同步更新所有歷史歸檔檔案的頂部導覽列，確保連結能往返切換
        import re
        if 'history_files' in locals() and history_files:
            print("-> 正在同步更新所有歷史歸檔網頁的導覽列...")
            for f_path in history_files:
                # 排除剛寫入的 archive_path，避免重複讀寫
                if os.path.abspath(f_path) == os.path.abspath(archive_path):
                    continue
                try:
                    with open(f_path, "r", encoding="utf-8") as f_old:
                        content_old = f_old.read()
                    
                    # 用最新產生的 history_html 替換舊有的 history-nav 導覽列
                    new_content_old = re.sub(
                        r'<div class="history-nav">.*?</div>',
                        history_html,
                        content_old,
                        flags=re.DOTALL
                    )
                    
                    # 順便更新每個舊頁面的 week-pill active 狀態，確保使用者清楚當前停留在哪一頁
                    basename_no_ext = os.path.basename(f_path).replace(".html", "")
                    
                    # 先將所有 pill 降為普通樣式，再把對應此舊週次檔案的按鈕設為 active
                    new_content_old = new_content_old.replace('week-pill active', 'week-pill')
                    new_content_old = new_content_old.replace(
                        f'href="{basename_no_ext}.html" class="week-pill"',
                        f'href="{basename_no_ext}.html" class="week-pill active"'
                    )
                    
                    with open(f_path, "w", encoding="utf-8") as f_old:
                        f_old.write(new_content_old)
                except Exception as ex:
                    print(f"  [警訊] 同步更新歷史檔案 {os.path.basename(f_path)} 導覽列失敗: {ex}")
    except Exception as e:
        print(f"  [警訊] 生成 HTML 看板時出錯: {e}")

def auto_deploy_to_github():
    """
    使用 subprocess 自動將最新的 HTML 網頁推送到 GitHub Pages。
    自動尋找包含 .git 的工作區根目錄。
    """
    import subprocess
    # 尋找包含 .git 的目錄（優先檢查 NEWS 目錄，若無則檢查上層目錄）
    base_dir = r"D:\ASUS\Anti-NotebookLM\NEWS"
    repo_dir = base_dir
    if not os.path.exists(os.path.join(base_dir, ".git")):
        parent_dir = os.path.dirname(base_dir)
        if os.path.exists(os.path.join(parent_dir, ".git")):
            repo_dir = parent_dir
            
    print("\n" + "=" * 60)
    print(" 🚀 啟動一鍵自動發布至 GitHub Pages...")
    print(f"  [工作目錄] {repo_dir}")
    print("=" * 60)
    
    try:
        # 1. git add .
        res_add = subprocess.run(["git", "add", "."], cwd=repo_dir, capture_output=True, encoding='utf-8', errors='replace')
        if res_add.returncode != 0:
            err_msg = (res_add.stderr or "").strip()
            print(f"  [警訊] Git add 失敗: {err_msg}")
            return False

        # 2. git commit -m
        commit_msg = f"Auto update news dashboard ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
        res_commit = subprocess.run(["git", "commit", "-m", commit_msg], cwd=repo_dir, capture_output=True, encoding='utf-8', errors='replace')
        
        stdout_str = res_commit.stdout or ""
        stderr_str = res_commit.stderr or ""
        
        if "nothing to commit" in stdout_str or "nothing to commit" in stderr_str:
            print("  [資訊] 本次變更內容無差異或已經是最新狀態。")
        elif res_commit.returncode != 0:
            print(f"  [資訊] Git commit 狀態: {stdout_str.strip() or stderr_str.strip()}")
            
        # 3. git push
        print("  -> 正在推送到雲端 GitHub...")
        res_push = subprocess.run(["git", "push"], cwd=repo_dir, capture_output=True, encoding='utf-8', errors='replace')
        if res_push.returncode != 0:
            # 嘗試設定 upstream 並推送到 origin main
            res_push = subprocess.run(["git", "push", "-u", "origin", "main"], cwd=repo_dir, capture_output=True, encoding='utf-8', errors='replace')
            if res_push.returncode != 0 and ("rejected" in res_push.stderr or "fetch first" in res_push.stderr):
                print("  [資訊] 偵測到雲端倉庫有舊歷史紀錄，正在進行首次強制對齊發布...")
                res_push = subprocess.run(["git", "push", "-u", "origin", "main", "--force"], cwd=repo_dir, capture_output=True, encoding='utf-8', errors='replace')
            
        if res_push.returncode == 0:
            print("  [成功] 網頁已成功自動推送到 GitHub！線上網址將在幾秒內更換為最新內容。")
            return True
        else:
            err_msg = (res_push.stderr or "").strip()
            print(f"  [警訊] Git push 失敗。請確認本機是否已連結 GitHub 倉庫。")
            print(f"  [錯誤細節] {err_msg}")
            return False
    except Exception as e:
        print(f"  [警訊] 自動發布過程發生異常: {e}")
        return False


def send_notification(success, title, message):
    """
    透過 .env 中設定的 Slack Incoming Webhook 及/或 Email SMTP 發送執行結果通知。
    兩種通知管道皆為選填，只有在對應環境變數存在時才會真正發送；都沒設定時此函式安靜跳過，不影響主流程。
    注意：LINE Notify 服務已於 2025/3/31 由官方終止服務，故不提供該管道，若需要 LINE 通知請改用
    LINE Messaging API (需另外申請 Bot 頻道，設定較複雜，故此處未實作)。
    """
    icon = "✅" if success else "❌"
    full_text = f"{icon} {title}\n{message}"

    slack_webhook = os.environ.get("SLACK_WEBHOOK_URL")
    if slack_webhook:
        try:
            requests.post(slack_webhook, json={"text": full_text}, timeout=10)
            print("  [資訊] 已發送 Slack 通知。")
        except Exception as e:
            print(f"  [警訊] Slack 通知發送失敗: {e}")

    smtp_host = os.environ.get("SMTP_HOST")
    notify_email_to = os.environ.get("NOTIFY_EMAIL_TO")
    if smtp_host and notify_email_to:
        try:
            import smtplib
            from email.mime.text import MIMEText

            smtp_port = int(os.environ.get("SMTP_PORT", "587"))
            smtp_user = os.environ.get("SMTP_USER")
            smtp_password = os.environ.get("SMTP_PASSWORD")

            msg = MIMEText(full_text)
            msg["Subject"] = title
            msg["From"] = smtp_user or notify_email_to
            msg["To"] = notify_email_to

            with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
                server.starttls()
                if smtp_user and smtp_password:
                    server.login(smtp_user, smtp_password)
                server.send_message(msg)
            print("  [資訊] 已發送 Email 通知。")
        except Exception as e:
            print(f"  [警訊] Email 通知發送失敗: {e}")


def append_to_excel(file_path, new_rows):
    """
    使用 openpyxl 將整理完的資料安全地寫入 Excel 的最上方，保留其它 Sheet 的內容與格式。
    會自動檢查並動態插入「對於筆電市場/華碩的影響」欄位，維持與現有格式的相容。
    在寫入被 Excel 鎖定時提供互動式提示。
    """
    # 完美對接的全新 headers (將新欄位設定在第九欄 I 欄，對應原本的 Unnamed: 8)
    headers = ['Year', 'Week', 'Category', 'Topic', 'Content', 'Link', '時間', '完整內容', '對於筆電市場/華碩的影響', 'Unnamed: 9']
    
    # 檢查 Excel 目錄是否存在，若不存在則建立
    dir_path = os.path.dirname(file_path)
    if dir_path and not os.path.exists(dir_path):
        os.makedirs(dir_path, exist_ok=True)

    if not os.path.exists(file_path):
        print(f"Excel 檔案不存在，正在全新建立：{file_path}")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "News"
        ws.append(headers)
    else:
        wb = openpyxl.load_workbook(file_path)
        if "News" not in wb.sheetnames:
            ws = wb.create_sheet("News")
            ws.append(headers)
        else:
            ws = wb["News"]
            
            # 動態檢測：檢查第 9 欄 (I 欄) 是否為「對於筆電市場/華碩的影響」
            # 如果不是 (原本是空的或 Unnamed: 8)，直接設定第 9 欄第一行為該 Header
            if ws.cell(row=1, column=9).value != '對於筆電市場/華碩的影響':
                print("  [資訊] 偵測到 Excel 第 9 欄 (I 欄) 尚未命名，將其命名為「對於筆電市場/華碩的影響」...")
                ws.cell(row=1, column=9, value='對於筆電市場/華碩的影響')
            
    # 將新資料插入在工作表的最前面（Header 正下方的第二行）
    num_new_rows = len(new_rows)
    if num_new_rows > 0:
        ws.insert_rows(2, num_new_rows)
        
        # 從第二行開始，依次填入新抓取的新聞
        for idx, row_data in enumerate(new_rows):
            current_row = 2 + idx
            for col_idx, h in enumerate(headers, 1):
                if h.startswith("Unnamed:"):
                    val = ""
                else:
                    val = row_data.get(h, "")
                ws.cell(row=current_row, column=col_idx, value=val)
        
    # 儲存工作簿並加入 PermissionError (檔案被 Excel 開啟中) 處理
    try:
        wb.save(file_path)
        print(f"\n[成功] 成功將 {len(new_rows)} 筆新聞資料寫入 Excel 檔案！")
    except PermissionError:
        print("\n" + "!"*60)
        print("[警告] 無法儲存 Excel 檔案！該檔案可能已被微軟 Excel 軟體開啟。")
        print("!"*60 + "\n")
        if AUTO_MODE:
            # 無人值守模式：不等待輸入，改為間隔重試數次後放棄
            retry_limit = 3
            for attempt in range(1, retry_limit + 1):
                print(f"  [無人值守模式] 等待 30 秒後重試存檔 (第 {attempt}/{retry_limit} 次)...")
                time.sleep(30)
                try:
                    wb.save(file_path)
                    print("[成功] 存檔成功！")
                    break
                except PermissionError:
                    continue
            else:
                print("[失敗] 多次重試後仍無法存檔，本次抓取的新聞資料未寫入 Excel，請手動確認檔案是否被佔用。")
        else:
            while True:
                ans = input("請先關閉該 Excel 檔案，然後輸入 'Y' 以重新嘗試存檔 (或輸入 'N' 放棄儲存): ")
                if ans.upper() == 'Y':
                    try:
                        wb.save(file_path)
                        print("[成功] 存檔成功！")
                        break
                    except PermissionError:
                        print("檔案仍被 Excel 鎖定，請關閉後再試。")
                elif ans.upper() == 'N':
                    print("[取消] 放棄儲存本次抓取的新聞資料。")
                    break

def process_category(category_name, query_base, current_year, current_week, existing_titles, existing_links, dedup_lock, dgt_session):
    """
    處理單一分類：搜尋候選新聞、逐篇解碼/抓取內文/AI 分析，回傳該分類收集到的 row_data 清單。
    此函式會被多個執行緒同時呼叫 (每個分類一個執行緒)，existing_titles/existing_links 是跨分類
    共用的去重集合，所有讀取後緊接寫入的檢查都放在 dedup_lock 內，確保「查到沒有 → 佔位」是原子操作，
    避免不同分類的執行緒同時抓到同一篇新聞卻都通過重複檢查。
    """
    print(f"\n-> 正在搜尋主題：【{category_name}】")
    category_collected_data = []

    all_candidates = []
    if category_name == "國際經濟" and GOOGLE_ALERT_RSS_ECON:
        print(f"  [{category_name}] [資訊] 偵測到 GOOGLE_ALERT_RSS_ECON，改從 Google 快訊 RSS 抓取新聞...")
        # 抓取最多 15 篇較多候選，以便挑選包含優先關鍵字的新聞
        all_candidates = fetch_google_alert_rss(GOOGLE_ALERT_RSS_ECON, max_results=15)

        # 優先權排序邏輯：優先處理「中國」、「美國」與宏觀的「全球經濟」，並降低「台灣本土企業/台股」新聞的優先度
        top_priority = ["中國", "美國", "美中", "聯準會", "Fed", "降息", "通膨"]
        macro_priority = ["全球經濟", "全球市場", "國際經濟", "全球成長"]
        exclude_keywords = ["董座", "董事長", "台股", "台廠", "三陽", "光陽", "台積電", "聯電", "鴻海"]

        def get_priority(item):
            title = item.get("title", "")
            # 1. 包含台灣本土企業/台股相關詞彙，優先度降到最低 (回傳 3)
            if any(x in title for x in exclude_keywords):
                return 3
            # 2. 包含美中等關鍵字，頂級優先 (回傳 0)
            if any(k in title for k in top_priority):
                return 0
            # 3. 包含宏觀全球經濟關鍵字，次級優先 (回傳 1)
            if any(k in title for k in macro_priority):
                return 1
            # 4. 包含普通「全球」字眼，三級優先 (回傳 2)
            if "全球" in title:
                return 2
            # 5. 其餘新聞 (回傳 3)
            return 3

        all_candidates.sort(key=get_priority)
        print(f"  [{category_name}] [資訊] 已根據優先關鍵字 {top_priority} 與排除詞彙重新排序候選新聞順序...")
    else:
        # 先找 DIGITIMES (優先，最多取2篇以保留空間給其他來源)
        digitimes_query = f"{query_base} site:digitimes.com.tw"
        digitimes_items = fetch_news_with_fallback(digitimes_query, max_results=2)

        # 額外訂閱官方 RSS 來源 (Adapter)：官方來源雜訊少、時間準確，優先於 Google 搜尋補充
        adapter_items = []
        for source in SOURCE_ADAPTERS.get(category_name, []):
            print(f"  [{category_name}] [資訊] 額外訂閱官方來源：{source['name']}...")
            adapter_items.extend(fetch_adapter_rss(source))

        # 搜尋其他管道補充
        print(f"  [{category_name}] [資訊] 搜尋其他管道新聞...")
        other_query = f"{query_base} -site:digitimes.com.tw"
        other_items = fetch_news_with_fallback(other_query, max_results=5)

        # 組合候選清單：DIGITIMES 優先，接著官方 RSS 來源，最後才是 Google News 搜尋補充
        all_candidates = digitimes_items + adapter_items + other_items

    if not all_candidates:
        print(f"  [{category_name}] -> 過去 3 天內未搜尋到相關新聞。")
        return category_collected_data

    print(f"  [{category_name}] -> 找到 {len(all_candidates)} 篇候選新聞，開始進行抓取與篩選...")

    valid_count = 0

    for idx, item in enumerate(all_candidates, 1):
        if valid_count >= MAX_ARTICLES_PER_CATEGORY:
            break

        title = item["title"]
        google_url = item["link"]
        pub_date = item["pub_date"]

        print(f"  [{category_name} {idx}/{len(all_candidates)}] 處理新聞：{title[:40]}...")

        with dedup_lock:
            if title.strip() in existing_titles:
                print(f"  [{category_name}] [跳過] 這篇新聞之前已經抓取過了 (標題重複)")
                continue
            # 立即佔位，避免其他分類的執行緒在本篇處理期間重複搶進同一篇新聞
            existing_titles.add(title.strip())

        # 1. 解碼 Google News 轉址 (若為官方 RSS Adapter 直接提供的真實連結，則略過解碼步驟)
        if item.get("direct"):
            real_url = google_url
        else:
            real_url = decode_url(google_url)
        cleaned_real_url = clean_url(real_url)

        with dedup_lock:
            if cleaned_real_url in existing_links:
                print(f"  [{category_name}] [跳過] 這篇新聞之前已經抓取過了 (網址重複: {cleaned_real_url})")
                continue
            existing_links.add(cleaned_real_url)

        # 過濾購物/零售/產品頁面（非新聞）與列表頁面
        _BLOCKED_DOMAINS = [
            'tw.buy.yahoo.com', 'buy.yahoo.com', 'shopping.pchome.com.tw',
            'momoshop.com.tw', 'shopee.tw', 'momo.dm', 'ecshop',
        ]
        _BLOCKED_URL_PATTERNS = ['/product/', '/products/', '/item/', '/goods/', 'goods.ruten']
        _BLOCKED_PAGE_PATTERNS = ['/category/', '/page/', '/search/', '/tags/', '/archive/']

        if any(d in real_url for d in _BLOCKED_DOMAINS) or any(p in real_url for p in _BLOCKED_URL_PATTERNS):
            print(f"  [{category_name}] [跳過] 疑似購物/產品頁面，略過 ({real_url[:70]}...)")
            continue

        if any(p in real_url.lower() for p in _BLOCKED_PAGE_PATTERNS):
            print(f"  [{category_name}] [跳過] 偵測為列表/分類分頁，非正文網頁 ({real_url[:70]}...)")
            continue

        # 2. 爬取內文純文字 (如果為 DIGITIMES 新聞會自動使用登入後的 session)
        web_text = extract_webpage_text(real_url, session=dgt_session)

        if not web_text or len(web_text.strip()) < 100:
            # 網頁抓取失敗或內文過短 (常見於 403 封鎖、JS 動態渲染頁面)：
            # 改用新聞標題結合 Gemini 知識庫進行背景分析，而非直接放棄整篇新聞
            print(f"  [{category_name}] [警訊] 無法取得完整網頁內文，改用標題進行 AI 背景分析 ({title[:30]}...)")
            title_analysis = analyze_news_from_title(title, category_name)
            time.sleep(1)

            if not title_analysis:
                print(f"  [{category_name}] [跳過] 標題分析亦失敗，略過此篇新聞 ({title[:30]}...)")
                continue

            raw_fallback = title.split(" - ")[0] if " - " in title else title
            topic = title_analysis.title.strip() if title_analysis.title and title_analysis.title.strip() else raw_fallback
            content = title_analysis.summary
            impact = title_analysis.impact_analysis
            pub_time = pub_date  # 無網頁內文可提取時間，改用 RSS 提供的原始發布時間
            full_content = ""

            # 強制進行新聞年份過濾，限制必須是當前執行年份 (2026)
            import re as _re
            year_match = _re.search(r'\b(20\d{2})\b', pub_time)
            if year_match and int(year_match.group(1)) < current_year:
                print(f"  [{category_name}] [跳過] 偵測為過期歷史舊聞，發布年份為 {year_match.group(1)} 年 ({topic[:30]}...)")
                continue
        else:
            # 3. 呼叫 Gemini 進行整理與摘要
            analysis = analyze_news_content(web_text)
            time.sleep(1)

            if analysis:
                # 優先使用 AI 提煉的精簡標題，若空則 fallback 到原始標題
                raw_fallback = title.split(" - ")[0] if " - " in title else title
                topic = analysis.title.strip() if analysis.title and analysis.title.strip() else raw_fallback
                content = analysis.summary
                impact = analysis.impact_analysis
                pub_time = analysis.publish_time
                full_content = analysis.content_clean

                # 強制進行新聞年份過濾，限制必須是當前執行年份 (2026)
                import re as _re
                year_match = _re.search(r'\b(20\d{2})\b', pub_time)
                if year_match:
                    extracted_year = int(year_match.group(1))
                    if extracted_year < current_year:
                        print(f"  [{category_name}] [跳過] 偵測為過期歷史舊聞，發布年份為 {extracted_year} 年 ({topic[:30]}...)")
                        continue
            else:
                print(f"  [{category_name}] [跳過] AI 分析失敗，略過此篇新聞 ({title[:30]}...)")
                continue

        # 過濾 UUID 式標題 (含連字符十六進位碼，如 0C813AA6-A7DD-41EA-975C)
        import re as _re
        if _re.search(r'[0-9A-Fa-f]{6,}-[0-9A-Fa-f]{4,}-[0-9A-Fa-f]{4,}', topic):
            print(f"  [{category_name}] [跳過] 標題為系統編碼，非正常新聞 ({topic[:50]})")
            continue
        # 過濾產品規格式標題 (含記憶體/儲存規格，如 8G+16G/2TB SSD/Win11)
        if _re.search(r'(?:/\d+G[B]?|\d+GB|SSD|Win11|Win10|PCIe\s|DDR\d)', topic, _re.IGNORECASE):
            print(f"  [{category_name}] [跳過] 標題含產品規格，疑似產品頁面 ({topic[:50]})")
            continue

        # 整理為 Excel 列結構
        row_data = {
            "Year": current_year,
            "Week": current_week,
            "Category": category_name,
            "Topic": topic,
            "Content": content,
            "對於筆電市場/華碩的影響": impact,
            "Link": real_url,
            "時間": pub_time,
            "完整內容": full_content
        }
        category_collected_data.append(row_data)
        valid_count += 1

    return category_collected_data

def main():
    print("=" * 60)
    print(" 每日新聞閱讀與 Excel 整理工具啟動")
    print(f" 執行時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f" Excel 目標：{EXCEL_PATH}")
    print("=" * 60)
    
    # 嘗試建立 DIGITIMES 會員 Session
    dgt_session = None
    if DIGITIMES_USER and DIGITIMES_PASSWORD:
        dgt_session = get_digitimes_session(DIGITIMES_USER, DIGITIMES_PASSWORD)
        
    # 計算當前的 ISO 年份與週數：一律採用「今天實際的 ISO 週次」，不再從 Excel 歷史紀錄推算 +1。
    # (若改成「歷史最新週次 +1」，一旦每天執行就會每天多跳一週，跟真實日曆完全對不上；
    #  改用實際 ISO 週次後，同一週內不管執行幾次都會正確歸入同一週報，只有真正跨週才會換週。)
    now = datetime.now()
    iso_year, iso_week, _ = now.isocalendar()
    current_year = iso_year
    current_week = iso_week

    # 嘗試讀取現有 Excel 中的歷史新聞，建立去重清單
    existing_titles = set()
    existing_links = set()
    if os.path.exists(EXCEL_PATH):
        try:
            df_history = pd.read_excel(EXCEL_PATH)
            if 'Topic' in df_history.columns:
                existing_titles = set(df_history['Topic'].dropna().astype(str).str.strip())
            if 'Link' in df_history.columns:
                existing_links = set(clean_url(url) for url in df_history['Link'].dropna().astype(str) if url.strip())
            print(f"  [資訊] 成功載入歷史紀錄：包含 {len(existing_titles)} 篇已抓取新聞，將自動過濾重複。")

            if 'Week' in df_history.columns and 'Year' in df_history.columns:
                max_year_in_excel = int(df_history['Year'].dropna().astype(float).max())
                df_latest_yr = df_history[df_history['Year'].dropna().astype(float).astype(int) == max_year_in_excel]
                max_week_vals = df_latest_yr['Week'].dropna()
                if not max_week_vals.empty:
                    print(f"\n  [確認] 偵測到 Excel 最新資料為 {max_year_in_excel}wk{int(max_week_vals.astype(float).max()):02d}")
                    print(f"  [確認] 今天實際週次為：{current_year}wk{current_week:02d}（本次將採用此週次寫入）")
                    if not AUTO_MODE:
                        ans = input(f"  是否改寫入其他週次？（直接 Enter 使用今天的實際週次，或輸入數字如 27 覆蓋）：").strip()
                        if ans.isdigit():
                            current_week = int(ans)
                    print(f"  [資訊] 本次寫入週次：{current_year}wk{current_week:02d}")
        except Exception as e:
            print(f"  [警訊] 無法讀取 Excel 歷史紀錄進行去重: {e}")
            
    collected_data = []

    # 並行處理四大分類：每個分類各自搜尋/抓取/分析，彼此獨立，用 ThreadPoolExecutor 同時執行
    # 以縮短整體等待網路 I/O 與 Gemini API 回應的時間。跨分類共用的去重集合由 dedup_lock 保護。
    dedup_lock = threading.Lock()
    print(f"\n-> 開始並行處理 {len(CATEGORIES)} 個分類 (最多同時 4 個分類一起抓取)...")
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_category = {
            executor.submit(
                process_category, category_name, query_base,
                current_year, current_week, existing_titles, existing_links,
                dedup_lock, dgt_session
            ): category_name
            for category_name, query_base in CATEGORIES.items()
        }
        for future in as_completed(future_to_category):
            category_name = future_to_category[future]
            try:
                collected_data.extend(future.result())
            except Exception as e:
                print(f"  [警訊] 分類【{category_name}】處理時發生未預期例外: {e}")

    if collected_data:
        # 將資料追加寫入 Excel
        append_to_excel(EXCEL_PATH, collected_data)
        # 產生 HTML (輸出為 index.html 與 week_num.html)
        print("\n-> 正在產生 HTML 報表...")
        local_index_path = r"D:\ASUS\Anti-NotebookLM\NEWS\index.html"
        generate_html_dashboard(EXCEL_PATH, local_index_path)

        # 無人值守模式下不自動開啟瀏覽器視窗 (排程執行時沒有人在旁邊看)
        if not AUTO_MODE:
            try:
                webbrowser.open(local_index_path)
                print("\n[成功] 已在瀏覽器中自動為您開啟「ASUS 新聞情報看板」網頁！")
            except Exception as e:
                print(f"  [警訊] 自動開啟網頁時出錯: {e}")

        # 自動推送到 GitHub
        deploy_success = auto_deploy_to_github()

        if deploy_success:
            send_notification(
                True, "ASUS 新聞情報看板 - 執行成功",
                f"本次共收集到 {len(collected_data)} 篇新聞，已成功發布至：\n"
                f"https://AngelaHNWang.github.io/Weekly-News-Summary/"
            )
        else:
            send_notification(
                False, "ASUS 新聞情報看板 - GitHub 發布失敗",
                f"本次共收集到 {len(collected_data)} 篇新聞並已寫入 Excel，"
                f"但推送到 GitHub Pages 失敗，請檢查本機 Git 設定或手動執行 git push。"
            )
    else:
        print("\n[結束] 今日沒有收集到任何新聞資料。")
        send_notification(
            True, "ASUS 新聞情報看板 - 本次無新資料",
            "本次排程已正常執行完畢，但搜尋範圍內沒有找到符合條件的新新聞，未寫入 Excel。"
        )

    print("\n" + "=" * 60)
    print(" 程式執行完畢")
    print("=" * 60)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("\n[嚴重錯誤] 程式執行時發生未預期的例外，本次執行中止：")
        traceback.print_exc()
        send_notification(
            False, "ASUS 新聞情報看板 - 執行失敗",
            f"news_collector.py 執行時發生未預期例外：{e}\n"
            f"請檢查排程執行紀錄，或手動執行以下指令排查問題：\npython news_collector.py --auto"
        )
        sys.exit(1)
