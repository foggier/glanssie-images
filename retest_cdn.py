# -*- coding: utf-8 -*-
"""403 alan URL'leri yavaş ve tarayıcı UA'sı ile yeniden dene."""
import json, os, time, urllib.request
TMP = r"C:\Users\drero\AppData\Local\hermes\cache\scratch"
d = json.load(open(os.path.join(TMP, "cdn_validation.json"), encoding="utf-8"))
base = "https://cdn.jsdelivr.net/gh/foggier/glanssie-images@main/"
bad = [b[1] for b in d["bad"]]
print("Yeniden denenecek:", len(bad))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
ok = 0; fail = []
for i, rel in enumerate(bad, 1):
    url = base + rel
    for attempt in (1, 2, 3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "image/*,*/*"})
            with urllib.request.urlopen(req, timeout=30) as r:
                data = r.read()
                if r.status == 200 and len(data) > 10000:
                    ok += 1
                    print(f"  [{i:2d}] OK {len(data)/1024:6.0f} KB  {rel}")
                    break
                raise RuntimeError(f"status {r.status}")
        except Exception as e:
            if attempt == 3:
                fail.append((rel, str(e)[:80]))
                print(f"  [{i:2d}] HATA {rel}: {str(e)[:60]}")
            else:
                time.sleep(4 * attempt)
    time.sleep(1.5)
print(f"\nSonuç: düzelen={ok}/{len(bad)}, kalıcı hata={len(fail)}")
for f in fail: print("   ", f)
json.dump({"retested": len(bad), "fixed": ok, "still_failing": fail},
          open(os.path.join(TMP, "cdn_retest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
