import sys
html_path = r'D:\ASUS\Anti-NotebookLM\NEWS\news_dashboard.html'
code_path = r'D:\ASUS\Anti-NotebookLM\NEWS\news_collector.py'

with open(code_path, 'r', encoding='utf-8') as f:
    code = f.read()

style_start = code.find('<style>')
style_end = code.find('</style>') + 8
new_style = code[style_start:style_end].replace('{{', '{').replace('}}', '}')

with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

old_style_start = html.find('<style>')
old_style_end = html.find('</style>') + 8
new_html = html[:old_style_start] + new_style + html[old_style_end:]

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(new_html)

print("CSS 已成功熱更新至 news_dashboard.html！")
