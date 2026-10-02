import urllib.request
import urllib.parse
import json

def test():
    text = urllib.parse.quote("руководитель отдела продаж")
    url = f"https://api.hh.ru/vacancies?text={text}&per_page=5"
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json"
    })
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            print("Found:", data.get("found"))
            for item in data.get("items", [])[:3]:
                print(item.get("name"), item.get("employer", {}).get("name"))
    except urllib.error.HTTPError as e:
        print("HTTP Error:", e.code, e.read().decode("utf-8"))
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    test()
