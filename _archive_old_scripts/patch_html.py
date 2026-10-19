import re

html_path = 'news_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

# Remove the subtitle
html = re.sub(r'<p class="subtitle">.*?</p>', '', html)

# Remove the icons from category titles
html = re.sub(r'<h3 class="cat-title">(📈|💻|⚙️|🏷️)\s*', '<h3 class="cat-title">', html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("HTML cleaned successfully!")
