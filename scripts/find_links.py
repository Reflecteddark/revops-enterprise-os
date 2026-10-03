import re

with open("docs/index.html", encoding="utf-8") as f:
    text = f.read()

urls = re.findall(r'https?://[^\s"\'<>]+', text)
for u in set(urls):
    if any(k in u for k in ['t.me', 'max', 'bot', 'dm1918', 'ai-rop']):
        print(u)
