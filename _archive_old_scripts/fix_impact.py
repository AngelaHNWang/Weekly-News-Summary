import re

html_path = 'news_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

# Pattern to match the old section-impact structure:
# <div class="section-impact">
#     <div class="section-impact-header">
#         <span>💡 AI觀點總結</span>
#     </div>
#     <div id="impact_section_..." class="impact-panel">
#         <p>[AI觀點] ...</p>
#     </div>
# </div>

pattern = re.compile(
    r'<div class="section-impact">\s*<div class="section-impact-header">\s*<span>💡 AI觀點總結</span>\s*</div>\s*<div id="(.*?)" class="impact-panel">\s*<p>(?:\[AI觀點\]\s*)?(.*?)</p>\s*</div>\s*</div>',
    re.DOTALL
)

def replacer(match):
    impact_id = match.group(1)
    content = match.group(2).strip()
    # The content might have <br> tags in it, that's fine.
    
    new_html = f'''<div class="section-impact" id="{impact_id}">
                        <p>💡 <strong>[AI觀點]</strong> {content}</p>
                    </div>'''
    return new_html

html = pattern.sub(replacer, html)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("HTML impact sections merged and highlighted successfully!")
