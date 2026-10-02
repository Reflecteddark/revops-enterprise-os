import urllib.request
import re

req = urllib.request.Request('https://business.max.ru', headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req).read().decode('utf-8', errors='ignore')
    for m in re.finditer(r'href="([^"]+)"', html):
        link = m.group(1)
        if any(w in link for w in ['doc', 'api', 'help', 't.me', 'max']):
            print('Found link:', link)
except Exception as e:
    print('Error:', e)
