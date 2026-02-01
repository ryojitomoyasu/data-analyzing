import re
import os

src = r"c:/Users/ryoji/OneDrive/GSF/00. Antigravity/戦略考案・相談アプリ2026/pages/1_📊_Sales_Dashboard.py"
dst = r"c:/Users/ryoji/OneDrive/GSF/00. Antigravity/戦略考案・相談アプリ2026/pages/2_📈_Sales_Comparison.py"

print(f"Reading from {src}")
with open(src, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Replace simple double quotes: key="val" -> key="val_comp"
content = re.sub(r'key="([^"]+)"', r'key="\1_comp"', content)

# 2. Replace simple single quotes: key='val' -> key='val_comp'
content = re.sub(r"key='([^']+)'", r"key='\1_comp'", content)

# 3. Replace f-string double: key=f"val_{x}" -> key=f"val_{x}_comp"
content = re.sub(r'key=f"([^"]+)"', r'key=f"\1_comp"', content)

# 4. Replace f-string single: key=f'val_{x}' -> key=f'val_{x}_comp'
content = re.sub(r"key=f'([^']+)'", r"key=f'\1_comp'", content)

print(f"Writing to {dst}")
with open(dst, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done.")
