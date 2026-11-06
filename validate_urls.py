"""Tüm CDN URL'lerini doğrula (jsDelivr)."""
import json, os, sys
from concurrent.futures import ThreadPoolExecutor
import urllib.request

TMP = r"C:\Users\drero\AppData\Local\hermes\cache\scratch"
MAP = os.path.join(r"C:\Users\drero\glanssie-cdn", "_mapping.json")
BASE = "https://cdn.jsdelivr.net/gh/foggier/glanssie-images@main/"

mapping = json.load(open(MAP, encoding="utf-8"))
urls = [(sku, rel) for sku, rels in mapping.items() for rel in rels]
print("Kontrol edilecek URL:", len(urls), flush=True)

def check(item):
    sku, rel = item
    url = BASE + rel
    try:
        req = urllib.request.Request(url, method="GET", headers={"User-Agent": "curl/8"})
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
            return (sku, rel, r.status, len(data), r.headers.get("Content-Type", ""))
    except Exception as e:
        return (sku, rel, getattr(e, "code", "ERR"), 0, str(e)[:60])

bad = []
ok = 0
with ThreadPoolExecutor(max_workers=6) as ex:
    for i, res in enumerate(ex.map(check, urls), 1):
        sku, rel, status, size, ct = res
        if status == 200 and size > 10000 and "image" in ct:
            ok += 1
        else:
            bad.append(res)
        if i % 50 == 0:
            print(f"  ...{i}/{len(urls)} (ok={ok}, hatalı={len(bad)})", flush=True)

print("\n=== SONUÇ ===")
print(f"Başarılı: {ok}/{len(urls)}")
print(f"Hatalı: {len(bad)}")
for b in bad[:30]:
    print("   ", b)
json.dump({"ok": ok, "total": len(urls), "bad": bad}, open(os.path.join(TMP, "cdn_validation.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
