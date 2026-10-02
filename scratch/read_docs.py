import urllib.request, re, ssl

ctx = ssl._create_unverified_context()
req = urllib.request.Request('https://dev.max.ru/docs/chatbots/bots-coding/examples', headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req, context=ctx).read().decode('utf-8', errors='ignore')
    text = re.sub('<[^<]+?>', ' ', html)
    with open('scratch/max_examples.txt', 'w', encoding='utf-8') as f:
        f.write(text)
    print("Saved scratch/max_examples.txt")
except Exception as e:
    print(e)
