# -*- coding: utf-8 -*-
"""Tüm görsel URL'lerini GitHub Pages üzerinden doğrula (rate-limit yok)."""
import json, os, sys, urllib.request
from concurrent.futures import ThreadPoolExecutor

BASE = sys.argv[1] if len(sys.argv) > 1 else "https://foggier.github.io/glanssie-images/"
MAP = r"C:\Users\drero\glanssie-cdn\_mapping.json"
TMP = r"C:\Users\drero\AppData\Local\hermes\cache\scratch"

mapping = json.load(open(MAP, encoding="utf-8"))
urls = [(sku, rel) for sku, rels in mapping.items() for rel in rels]
print("Kontrol edilecek:", len(urls), "| host:", BASE, flush=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"


def check(item):
    sku, rel = item
    try:
        req = urllib.request.Request(BASE + rel, headers={"User-Agent": UA, "Accept": "image/*"})
        with urllib.request.urlopen(req, timeout=45) as r:
            data = r.read()
            return (sku, rel, r.status, len(data), r.headers.get("Content-Type", ""))
    except Exception as e:
        return (sku, rel, getattr(e, "code", "ERR"), 0, str(e)[:70])


bad, ok = [], 0
with ThreadPoolExecutor(max_workers=10) as ex:
    for i, res in enumerate(ex.map(check, urls), 1):
        sku, rel, status, size, ct = res
        if status == 200 and size > 10000 and "image" in ct:
            ok += 1
        else:
            bad.append(res)
        if i % 100 == 0:
            print(f"  ...{i}/{len(urls)}  ok={ok} hatalı={len(bad)}", flush=True)

print(f"\n=== SONUÇ ({BASE}) ===")
print(f"Başarılı: {ok}/{len(urls)} | Hatalı: {len(bad)}")
for b in bad[:20]:
    print("   ", b)
json.dump({"base": BASE, "ok": ok, "total": len(urls), "bad": bad},
          open(os.path.join(TMP, "cdn_validation_pages.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
