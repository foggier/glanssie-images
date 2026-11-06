# -*- coding: utf-8 -*-
"""GLANSSIE CDN v2 — EXIF rotasyonu UYGULANARAK yeniden üretim.

SORUN: Kaynak görsellerin tamamı EXIF Orientation=6 (90 derece) taşıyor.
PIL pixel verisini döndürmez, yalnızca EXIF etiketini okur; etiket düşürülünce
görsel yan kalır. Çözüm: ImageOps.exif_transpose() ile döndürmeyi PIXEL'e uygula.
"""
import os, json, time
from PIL import Image, ImageOps

SRC  = r"C:\Users\drero\OneDrive\Masaüstü\GLANSSİE\YÜKLENECEK ÜRÜNLER"
REPO = r"C:\Users\drero\glanssie-cdn"
MAX_IMG, LONG, Q = 8, 1600, 85

# 230 numaralı klasör 17 ayrı ürüne bölündü: yeni kod -> kaynak görsel numaraları
CHARM = {
 'BDHA2025-245': ['02'], 'BDHA2025-246': ['03'], 'BDHA2025-247': ['04'],
 'BDHA2025-248': ['05', '06'], 'BDHA2025-249': ['07'], 'BDHA2025-250': ['08'],
 'BDHA2025-251': ['09'], 'BDHA2025-252': ['10'], 'BDHA2025-253': ['11'],
 'BDHA2025-254': ['12'], 'BDHA2025-255': ['13'], 'BDHA2025-256': ['14'],
 'BDHA2025-257': ['15'], 'BDHA2025-258': ['16'], 'BDHA2025-259': ['17'],
 'BDHA2025-260': ['18'], 'BDHA2025-261': ['19'],
}
CHARM_SRC = os.path.join(SRC, "BDHA2025-230")


def export(src_paths, outdir, sku):
    os.makedirs(outdir, exist_ok=True)
    rels = []
    for n, sp in enumerate(src_paths, 1):
        dst = os.path.join(outdir, f"{sku}-{n:02d}.jpg")
        with Image.open(sp) as im:
            im = ImageOps.exif_transpose(im)      # <-- KRİTİK: rotasyonu pixel'e uygula
            im = im.convert("RGB")
            im.thumbnail((LONG, LONG), Image.LANCZOS)
            im.save(dst, "JPEG", quality=Q, optimize=True, progressive=True)
        rels.append(f"{sku}/{sku}-{n:02d}.jpg")
    return rels


mapping = {}
t0 = time.time()

# 1) normal SKU'lar
skus = sorted(d for d in os.listdir(SRC) if os.path.isdir(os.path.join(SRC, d)) and d != "BDHA2025-230")
for i, sku in enumerate(skus, 1):
    files = sorted(os.listdir(os.path.join(SRC, sku)))[:MAX_IMG]
    mapping[sku] = export([os.path.join(SRC, sku, f) for f in files],
                          os.path.join(REPO, sku), sku)
    print(f"[{i:3d}/{len(skus)}] {sku}: {len(mapping[sku])} görsel", flush=True)

# 2) charm ürünleri
for sku, nums in CHARM.items():
    mapping[sku] = export([os.path.join(CHARM_SRC, f"italyan_bileklik_charm-{n}.jpg") for n in nums],
                          os.path.join(REPO, sku), sku)
    print(f"[charm] {sku}: {len(mapping[sku])} görsel", flush=True)

json.dump(mapping, open(os.path.join(REPO, "_mapping.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print(f"\nÜrün: {len(mapping)} | Görsel: {sum(len(v) for v in mapping.values())} | Süre: {time.time()-t0:.0f} sn")
