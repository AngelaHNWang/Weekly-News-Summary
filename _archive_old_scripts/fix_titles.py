import re

html_path = 'news_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

# The original titles from the News.xlsx were:
# 1. OECD下修全球經濟成長預測 國際景氣展望趨於保守
# 2. AI PC掀新戰，高通低價筆電爭市場，然筆電出貨仍疲弱
# 3. AI點燃上游戰火 NVIDIA挺進處理器 記憶體DRAM HBM需求噴發
# 4. 品牌廠以AI硬體及高階產品創新 驅動PC市場新需求與成長

original_titles = [
    "OECD下修全球經濟成長預測 國際景氣展望趨於保守",
    "AI PC掀新戰，高通低價筆電爭市場，然筆電出貨仍疲弱",
    "AI點燃上游戰火 NVIDIA挺進處理器 記憶體DRAM HBM需求噴發",
    "品牌廠以AI硬體及高階產品創新 驅動PC市場新需求與成長"
]

def format_title(text):
    # Replace Chinese punctuation with <br>
    text = re.sub(r'[，、。！：；]', '<br>', text)
    # Replace spaces with <br> UNLESS the space is between two alphanumeric characters
    text = re.sub(r'(?<![A-Za-z0-9])\s+|\s+(?![A-Za-z0-9])', '<br>', text)
    # Clean up any potential double <br> or trailing <br>
    text = re.sub(r'(<br>)+', '<br>', text).strip('<br>')
    return text

formatted_titles = [format_title(t) for t in original_titles]

# Replace the existing titles in the HTML
# Since the HTML currently has the "cleaned" versions without punctuation, 
# we can just use re.sub with a replacement function that pops from our formatted_titles list.
title_index = 0
def replacer(match):
    global title_index
    if title_index < len(formatted_titles):
        new_title = formatted_titles[title_index]
        title_index += 1
        return f'<h3 class="cat-title">{new_title}</h3>'
    return match.group(0)

html = re.sub(r'<h3 class="cat-title">.*?</h3>', replacer, html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("HTML titles fixed using smart <br> formatting!")
