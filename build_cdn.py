"""GLANSSIE - Trendyol CDN görsel optimizasyon hattı.
Kaynak: OneDrive GLANSSİE klasörü -> Hedef: C:\\Users\\drero\\glanssie-cdn (git repo)
Her SKU için en fazla 8 görsel, uzun kenar 1600px, JPEG q85, progressive.
Çıktı adı: <SKU>/<SKU>-NN.jpg
"""
import os, sys, json, io, time
from PIL import Image

SRC  = r"C:\Users\drero\OneDrive\Masaüstü\GLANSSİE\YÜKLENECEK ÜRÜNLER"
TMP  = r"C:\Users\drero\AppData\Local\hermes\cache\scratch"
DEST = r"C:\Users\drero\glanssie-cdn"
MAX_IMG = 8
LONG_SIDE = 1600
QUALITY = 85

inv = json.load(open(os.path.join(TMP, "glanssie_inventory.json"), encoding="utf-8"))
os.makedirs(DEST, exist_ok=True)

mapping = {}          # sku -> [cdn relative paths]
total_bytes = 0
count = 0
errors = []
t0 = time.time()

for i, r in enumerate(inv, 1):
    sku = r["sku"]
    outdir = os.path.join(DEST, sku)
    os.makedirs(outdir, exist_ok=True)
    files = r["files"][:MAX_IMG]
    rels = []
    for n, fn in enumerate(files, 1):
        src = os.path.join(SRC, sku, fn)
        out = os.path.join(outdir, f"{sku}-{n:02d}.jpg")
        try:
            with Image.open(src) as im:
                im = im.convert("RGB")
                im.thumbnail((LONG_SIDE, LONG_SIDE), Image.LANCZOS)
                im.save(out, "JPEG", quality=QUALITY, optimize=True, progressive=True)
            sz = os.path.getsize(out)
            total_bytes += sz
            count += 1
            rels.append(f"{sku}/{sku}-{n:02d}.jpg")
        except Exception as e:
            errors.append((sku, fn, str(e)))
    mapping[sku] = rels
    print(f"[{i:3d}/{len(inv)}] {sku}: {len(rels)} görsel -> {sum(os.path.getsize(os.path.join(DEST,x)) for x in rels)/1024:.0f} KB", flush=True)

json.dump(mapping, open(os.path.join(DEST, "_mapping.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n=== ÖZET ===")
print("Görsel sayısı:", count)
print("Toplam boyut: %.1f MB" % (total_bytes/1024/1024))
print("Hata:", len(errors))
for e in errors: print("   ", e)
print("Süre: %.1f sn" % (time.time()-t0))
