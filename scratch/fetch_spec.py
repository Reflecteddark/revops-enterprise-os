import urllib.request
import re
import ssl

ctx = ssl._create_unverified_context()
req = urllib.request.Request('https://dev.max.ru/docs/open-api', headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req, context=ctx).read().decode('utf-8', errors='ignore')
    urls = re.findall(r'https?://[^\s"\'<>]+\.yaml', html)
    print('Found YAML urls:', urls)
    if urls:
        raw = urllib.request.urlopen(urls[0], context=ctx).read().decode('utf-8', errors='ignore')
        with open('scratch/max_spec.yaml', 'w', encoding='utf-8') as f:
            f.write(raw)
        print('Saved scratch/max_spec.yaml:', len(raw))
except Exception as e:
    print('Error:', e)
