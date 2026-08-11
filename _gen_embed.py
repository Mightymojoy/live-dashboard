# -*- coding: utf-8 -*-
import json, os, re

src = r'D:\WK工作空间\live-dashboard\live-dashboard\trend-demo.html'
out_path = r'D:\WK工作空间\live-dashboard\live-dashboard\trend-demo-embed.html'
data_path = r'D:\WK工作空间\live-dashboard\live-dashboard\live_data.json'

with open(src, 'rb') as f:
    html = f.read().decode('utf-8')

with open(data_path, 'rb') as f:
    data = json.load(f)
data_obj = json.dumps(data, ensure_ascii=False)

# 替换 fetch 加载块为内嵌数据
pattern = re.compile(
    r"    // 数据加载\n    fetch\('\./live_data\.json'\)\n      \.then\(r => r\.json\(\)\)\n      \.then\(data => \{"
)

new_block = (
    "    // 数据加载（Demo内嵌数据，无需服务器）\n"
    "    const EMBEDDED_DATA = " + data_obj + ";\n"
    "    Promise.resolve(EMBEDDED_DATA)\n"
    "      .then(data => {"
)

html2, count = pattern.subn(new_block, html)
print(f'替换 fetch 块: {count} 处')

with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
    f.write(html2)

print(f'OK 生成 trend-demo-embed.html ({os.path.getsize(out_path)/1024/1024:.1f}MB)')
print('EMBEDDED_DATA 对象:', 'const EMBEDDED_DATA = {"' in html2)
print('lucide 图标数:', html2.count('data-lucide='))
print('排名medal:', html2.count('data-lucide="medal"'))
