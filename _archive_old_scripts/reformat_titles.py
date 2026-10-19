import re

html_path = 'news_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

def replace_punctuation(match):
    # Extract the title content
    title = match.group(1)
    # Replace space, full-width comma, half-width comma, etc with <br>
    formatted = re.sub(r'[\s，、。]+', '<br>', title.strip())
    return f'<h3 class="cat-title">{formatted}</h3>'

html = re.sub(r'<h3 class="cat-title">(.*?)</h3>', replace_punctuation, html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("HTML titles reformatted successfully!")
