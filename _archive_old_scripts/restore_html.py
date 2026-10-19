import re

html_path = 'news_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

def clean_title(match):
    title = match.group(1)
    # Replace <br> with space
    title = title.replace('<br>', ' ')
    # Remove all Chinese and English punctuation
    title = re.sub(r'[，、。！：；,.;!:]', ' ', title)
    # Collapse multiple spaces into one
    title = re.sub(r'\s+', ' ', title).strip()
    return f'<h3 class="cat-title">{title}</h3>'

html = re.sub(r'<h3 class="cat-title">(.*?)</h3>', clean_title, html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("HTML titles cleaned successfully!")
