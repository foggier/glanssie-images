# -*- coding: utf-8 -*-
"""GLANSSIE -> Trendyol ürün oluşturma şablonu doldurucu.

Girdi : glanssie_model.json (SKU, kategori, görsel URL'leri, stok, maliyet, nitelik)
        _mapping.json (SKU -> CDN göreli yollar)
Çıktı : urun-olusturma-1513739-DOLDURULMUS.xlsx

KURALLAR
 - Kolonlar SADECE başlık adıyla çözümlenir (indeks kullanılmaz).
 - Hiçbir nitelik değeri uydurulmaz; her değer o kategorinin enum sayfasında doğrulanır.
 - Fiyat kolonları bilinçli olarak BOŞ bırakılır (kullanıcı dolduracak).
 - Üretici/İthalatçı (regülasyon) kolonları kullanıcı bilgisi gelene kadar BOŞ.
 - Marka adı ürün başlığına YAZILMAZ (marka kolonu = glanssieraad).
"""
import os, json, re, collections, unicodedata
import openpyxl

TMP   = r"C:\Users\drero\AppData\Local\hermes\cache\scratch"
BASE  = r"C:\Users\drero\OneDrive\Masaüstü\GLANSSİE"
TPL   = os.path.join(BASE, "urun-olusturma-1513739.xlsx")
MAP   = r"C:\Users\drero\glanssie-cdn\_mapping.json"
OUT   = os.path.join(BASE, "urun-olusturma-1513739-DOLDURULMUS.xlsx")
CDN   = "https://foggier.github.io/glanssie-images/"

BRAND   = "glanssieraad"
KDV     = 20      # TODO: kullanıcı teyidi
DESI    = 0.5     # TODO: kullanıcı teyidi
LEAD    = 2       # sevkiyat süresi (gün)
MENGEI  = ""      # TODO: kullanıcı teyidi (TR / CN ...)
CARE    = "Ürünün kullanım ömrünü korumak için ürüne özel bakım talimatlarını takip ediniz."

SHEET_OF_CAT = {"Kolye": "Çelik Kolye(2853)", "Bileklik": "Çelik Bileklik(2845)",
                "Küpe": "Çelik Küpe(2846)", "Yüzük": "Çelik Yüzük(2841)"}
ENUM_SHEET = {"Kolye": "Urun_Ozellik_Bilgileri_2853", "Bileklik": "Urun_Ozellik_Bilgileri_2845",
              "Küpe": "Urun_Ozellik_Bilgileri_2846", "Yüzük": "Urun_Ozellik_Bilgileri_2841"}

# ------------------------------------------------------------------ nitelik sözlükleri
MOTIF = {
 'kalp':            ("kalp", "Kalp Motifli", "Kalp motifi, sevgiyi ve bağlılığı simgeleyen zamansız bir tasarım dilidir"),
 'kuzey_yildizi':   ("kuzey yıldızı", "Kuzey Yıldızı Motifli", "Kuzey yıldızı motifi, yol göstericiliği ve umudu temsil eder"),
 'yildiz':          ("yıldız", "Yıldız Motifli", "Yıldız motifi, ışıltıyı ve hayalleri simgeleyen klasik bir tercihtir"),
 'ay_ve_yildiz':    ("ay ve yıldız", "Ay ve Yıldız Figürlü", "Ay ve yıldız figürü, gökyüzü temalı zarif bir bütünlük sunar"),
 'ay':              ("ay", "Ay Figürlü", "Ay figürü, sakinliği ve zarafeti yansıtan ince bir detaydır"),
 'ayi':             ("ayıcık", "Ayıcık Figürlü", "Sevimli ayıcık figürü, eğlenceli ve sempatik bir görünüm kazandırır"),
 'gunes':           ("güneş", "Güneş Motifli", "Güneş motifi, enerjiyi ve pozitifliği simgeler"),
 'yonca':           ("yonca", "Yonca Motifli", "Yonca motifi, şansın ve iyi talihin simgesidir"),
 'kelebek':         ("kelebek", "Kelebek Figürlü", "Kelebek figürü, hafifliği, zarafeti ve dönüşümü anlatır"),
 'cift_kelebek':    ("çift kelebek", "Çift Kelebek Figürlü", "Çift kelebek figürü, simetrik ve romantik bir kompozisyon oluşturur"),
 'cicek':           ("çiçek", "Çiçek Figürlü", "Çiçek figürü, doğallığı ve tazeliği yansıtır"),
 'gul':             ("gül", "Gül Motifli", "Gül motifi, klasik romantizmin en bilinen sembolüdür"),
 'papatya':         ("papatya", "Papatya Motifli", "Papatya motifi, sadeliği ve doğal güzelliği simgeler"),
 'balik':           ("balık", "Balık Figürlü", "Balık figürü, deniz temalı özgün bir tasarım sunar"),
 'anka_kusu':       ("anka kuşu", "Anka Kuşu Motifli", "Anka kuşu motifi, yeniden doğuşu ve sonsuz yenilenmeyi simgeler"),
 'kugu':            ("kuğu", "Kuğu Figürlü", "Kuğu figürü, zarafeti ve asaleti temsil eder"),
 'yusufcuk':        ("yusufçuk", "Yusufçuk Figürlü", "Yusufçuk figürü, uğur ve doğal denge ile ilişkilendirilir"),
 'yengec':          ("yengeç", "Yengeç Figürlü", "Yengeç figürü, astrolojik temalı özgün bir detaydır"),
 'goz_ve_kiprik':   ("göz ve kirpik", "Göz ve Kirpik Figürlü", "Göz ve kirpik figürü, ifade gücü yüksek modern bir motif oluşturur"),
 'goz':             ("göz", "Göz Figürlü", "Göz figürü, nazara karşı koruyucu anlam taşır"),
 'kilic':           ("kılıç", "Kılıç Motifli", "Kılıç motifi, gücü ve keskin bir stili yansıtır"),
 'yelpaze':         ("yelpaze", "Yelpaze Motifli", "Yelpaze motifi, geometrik ve dengeli bir form sunar"),
 'spiral':          ("spiral", "Spiral Motifli", "Spiral motifi, sonsuzluğu ve akışkanlığı simgeler"),
 'hasir':           ("hasır", "Hasır Motifli", "Hasır motifi, dokulu ve bohem bir görünüm sağlar"),
 'orgu':            ("örgü", "Örgü Motifli", "Örgü motifi, iç içe geçen detaylarıyla hacimli bir görünüm sunar"),
 'kare':            ("kare", "Kare Motifli", "Kare motifli geometrik hatlar modern bir estetik oluşturur"),
 'yuvarlak':        ("yuvarlak", "Yuvarlak Motifli", "Yuvarlak motifli formlar yumuşak ve dengeli bir görünüm verir"),
 'serenay':         ("serenay", "Serenay Model", "Serenay modeli, trend odaklı özel bir tasarım çizgisine sahiptir"),
 'eyfel':           ("Eyfel", "Eyfel Temalı", "Eyfel teması, seyahat ve şehir tutkunlarına hitap eder"),
 'ucak':            ("uçak", "Uçak Figürlü", "Uçak figürü, seyahat ve özgürlük temasını taşır"),
 'gezgin':          ("gezgin", "Gezgin Temalı", "Gezgin teması, keşfetme isteğini simgeler"),
 'seyahat':         ("seyahat", "Seyahat Temalı", "Seyahat teması, anı biriktirme duygusunu yansıtır"),
 'okyanus':         ("okyanus", "Okyanus Temalı", "Okyanus teması, huzuru ve derinliği çağrıştırır"),
 'nazar':           ("nazar", "Nazar Temalı", "Nazar teması, geleneksel koruma inancını modern bir tasarıma taşır"),
 'nazar_boncugu':   ("nazar boncuğu", "Nazar Boncuğu Detaylı", "Nazar boncuğu detayı, koruyucu anlamıyla öne çıkar"),
 'fiyonk':          ("fiyonk", "Fiyonk Temalı", "Fiyonk teması, feminen ve romantik bir dokunuş katar"),
 'kar_tanesi':      ("kar tanesi", "Kar Tanesi Motifli", "Kar tanesi motifi, kış temalı kristal bir görünüm sunar"),
 'yaprak':          ("yaprak", "Yaprak Figürlü", "Yaprak figürü, doğayla uyumlu ince bir detaydır"),
 'ahtapot':         ("ahtapot", "Ahtapot Figürlü", "Ahtapot figürü, deniz altı temalı sıra dışı bir tasarımdır"),
 'zincir_ve_balik': ("zincir ve balık", "Zincir ve Balık Figürlü", "Zincir ve balık figürü, ikili kompozisyon oluşturur"),
 'rehber':          ("rehber", "Rehber Figürlü", "Rehber figürü, yolculuk temasını tamamlayan bir detaydır"),
 'kiraz':           ("kiraz", "Kiraz Motifli", "Kiraz motifi, tatlı ve canlı bir görünüm kazandırır"),
 'gezegen':         ("gezegen", "Gezegen Motifli", "Gezegen motifi, gökyüzü temasını zenginleştirir"),
 'kibrit':          ("kibrit", "Kibrit Çubuklu", "Kibrit çubuğu formu, uzun ve ince bir silüet oluşturur"),
 'sahmeran':        ("şahmeran", "Şahmeran Motifli", "Şahmeran motifi, Anadolu mitolojisinden gelen koruyucu bir simgedir"),
 'damla':           ("damla", "Damla Kesim", "Damla kesim form, su gibi akışkan ve zarif bir hat sunar"),
 'tek_tas':         ("tek taş", "Tek Taş", "Tek taş tasarım, sadeliğin en klasik ifadesidir"),
 'charm':           ("charm", "Charm Detaylı", "Charm detayları, kişiselleştirilebilir bir kullanım sağlar"),
 'su_yolu':         ("su yolu", "Su Yolu", "Su yolu dizilimi, kesintisiz bir ışıltı hattı oluşturur"),
 'love':            ("love", "Love Yazılı", "Love yazısı, duyguyu doğrudan anlatan bir detaydır"),
 'geometrik':       ("geometrik", "Geometrik", "Geometrik hatlar çağdaş ve iddialı bir görünüm sunar"),
 'cift_motifli':    ("çift motif", "Çift Motifli", "Çift motif, iki ayrı formun dengeli bir kompozisyonda birleştiği bir tasarım sunar"),
}

# Motif ÖNCELİK sırası: özgül motifler önce, genel temalar (seyahat/gezgin/okyanus) en sonda.
# Uzunluğa göre seçmek yanlıştı: 'seyahat' gibi genel bir tema, 'eyfel' gibi özgül motifi eziyordu.
MOTIF_PRIORITY = [
 'kalp', 'kuzey_yildizi', 'ay_ve_yildiz', 'yildiz', 'ay', 'ayi', 'gunes', 'yonca',
 'kelebek', 'cift_kelebek', 'cicek', 'gul', 'papatya', 'balik', 'anka_kusu', 'kugu',
 'yusufcuk', 'yengec', 'goz_ve_kiprik', 'goz', 'kilic', 'yelpaze', 'spiral', 'hasir',
 'orgu', 'kare', 'yuvarlak', 'eyfel', 'ucak', 'yaprak', 'ahtapot', 'zincir_ve_balik',
 'rehber', 'kiraz', 'gezegen', 'geometrik', 'sahmeran', 'kibrit', 'damla', 'tek_tas',
 'cift_motifli', 'charm', 'su_yolu', 'love', 'fiyonk', 'kar_tanesi', 'nazar_boncugu',
 'nazar', 'serenay', 'gezgin', 'seyahat', 'okyanus',
]
STONE = {
 'Zirkon':      ("Zirkon Taşlı", "zirkon taşları", "Işığı yansıtan zirkon taşları, ürüne canlı bir parlaklık kazandırır."),
 'İnci':        ("İnci Detaylı", "inci detayları", "İnci detayları, klasik ve sofistike bir görünüm sunar."),
 'Boncuk':      ("Boncuk Detaylı", "boncuk detayları", "Boncuk detayları, renkli ve hareketli bir doku oluşturur."),
}
COLOR = {"Beyaz": "beyaz", "Mavi": "mavi", "Pembe": "pembe", "Kırmızı": "kırmızı",
         "Yeşil": "yeşil", "Zümrüt Yeşili": "zümrüt yeşili", "Çok Renkli": "çok renkli",
         "Kahverengi": "kahverengi"}
TYPE = {"Kolye":    ("kolye", "Kolye, dekoltenizi tamamlayan ve kombininize zarafet katan bir aksesuardır.", "ayarlanabilir zincir yapısı"),
        "Bileklik": ("bileklik", "Bileklik, bileğinize zarif bir vurgu yapan günlük bir takıdır.", "pratik klipsli kilit mekanizması"),
        "Küpe":     ("küpe", "Küpe, yüz hatlarınızı öne çıkaran ve stilinizi tamamlayan bir takıdır.", "hafif yapısı ve güvenli kilit sistemi"),
        "Yüzük":    ("yüzük", "Yüzük, günlük kombinlerden özel davetlere uzanan geniş bir kullanım sunar.", "pürüzsüz iç yüzeyi")}

OPENINGS = [
 "Günlük kombinlerinize zarif bir dokunuş katmak için tasarlandı.",
 "Modern çizgileri ve ince işçiliğiyle kombininize kolayca uyum sağlar.",
 "Hem günlük kullanım hem de özel davetler için ideal bir tercihtir.",
 "Sade ama dikkat çekici duruşuyla stilinize karakter katar.",
 "Zarif detaylarıyla takı kutunuzun vazgeçilmez parçası olmaya adaydır.",
 "Kombinlerinizi tamamlayan şık ve kullanışlı bir aksesuar arayanlar için ideal.",
]
CLOSINGS = [
 "Hediye paketine uygun yapısıyla sevdiklerinize özel günlerde verebileceğiniz anlamlı bir seçenektir.",
 "Şık sunumu sayesinde doğum günü, yıl dönümü ve özel günlerde gönül rahatlığıyla hediye edilebilir.",
 "Kutusunda teslim edilir; hem kendiniz hem de sevdikleriniz için keyifli bir hediye alternatifidir.",
 "Zamanla değerini koruyan tasarımıyla uzun süre severek kullanabileceğiniz bir takıdır.",
 "Kombin esnekliği sayesinde hem spor hem şık kombinlerde rahatça tercih edilebilir.",
 "Hediye edilmeye uygun tasarımıyla özel günlerinizi güzelleştirir.",
]

def pick(lst, key, offset=0):
    return lst[(sum(ord(c) for c in key) + offset) % len(lst)]

def ascii_low(s):
    """Türkçe karakterleri sadeleştirip küçük harfe çevir (yalnızca çözümleme için)."""
    s = s.lower().replace('ı', 'i').replace('İ', 'i').replace('ş', 's').replace('ğ', 'g')
    s = s.replace('ü', 'u').replace('ö', 'o').replace('ç', 'c')
    return s


def compose_title(m):
    n = m["nitelik"]
    parts = []
    if n["metal"] == "Altın Kaplama": parts.append("Altın Kaplama")
    elif n["metal"] == "Gümüş":       parts.append("Gümüş Renkli")
    wc_t = n.get("web_color")
    if wc_t in COLOR: parts.append(wc_t)   # Altın/Gümüş/Metalik metal adıdır, renk olarak yazılmaz
    tr_t = n.get("tas_rengi")
    if tr_t and tr_t != wc_t: parts.append(tr_t)   # taş rengi başlıkta bilgi değeri taşır
    if n.get("mineli"): parts.append("Mineli")
    if n.get("stone_phrase"): parts.append(n["stone_phrase"])
    for key in n["motifler"]:
        if key in MOTIF:
            parts.append(MOTIF[key][1]); break
    if n.get("stil"): parts.append(n["stil"])
    parts.append("Çelik")
    parts.append(m["kategori"])
    if m.get("set_mi"): parts.append("Seti")
    t = " ".join(parts)
    t = re.sub(r"\s+", " ", t).strip()
    if len(t) > 100: t = t[:97].rstrip() + "..."
    return t

def compose_desc(m):
    n = m["nitelik"]; sku = m["sku"]
    col = n["renkler"][0] if n["renkler"] else None
    tname, tdesc, tfeat = TYPE[m["kategori"]]
    bits = []
    bits.append(pick(OPENINGS, sku))
    if n["metal"] == "Altın Kaplama":
        bits.append("Altın kaplama çelik gövdesi, günlük kullanım için ideal bir dayanıklılık sunar.")
    elif n["metal"] == "Gümüş":
        bits.append("Gümüş renkli çelik gövdesi, günlük kombinlerde rahatlıkla tercih edilebilir.")
    else:
        bits.append("Paslanmaz çelik gövdesi, hafif ve günlük kullanıma uygun bir yapı sunar.")
    for key in n["motifler"]:
        if key in MOTIF:
            bits.append(MOTIF[key][2] + "."); break
    if n.get("tas_cinsi_beyan") and n["tas_cinsi_beyan"] in STONE:
        bits.append(STONE[n["tas_cinsi_beyan"]][2])
    elif n["tasli"]:
        bits.append("Ürün üzerindeki taş detayları, ışığı yansıtan canlı bir görünüm oluşturur.")
    wc = n.get("web_color")
    if wc == "Çok Renkli":
        bits.append("Birden fazla renk tonunu bir arada sunan tasarımıyla kombinlere hareket katar.")
    elif wc in COLOR:
        bits.append(f"Ürünün hâkim rengi {COLOR[wc]} tonudur; bu ton günlük kombinlerle kolayca uyum sağlar.")
    tr = n.get("tas_rengi")
    if tr:
        bits.append(f"Üzerindeki {COLOR.get(tr, tr.lower())} tonlu taşlar tasarıma canlılık katar.")
    bits.append(f"{tdesc} {tfeat.capitalize()} sayesinde gün boyu konforlu bir kullanım sunar.")
    if m.get("set_mi"):
        bits.append("Set içeriği tek bir ürün olarak, birbiriyle uyumlu parçalar hâlinde sunulur.")
    bits.append(pick(CLOSINGS, sku, 3))
    d = " ".join(bits)
    d = re.sub(r"\s+", " ", d).strip()
    return d

# set tespiti
def is_set(name):
    """Yalnızca gerçek çok parçalı ürünler: '5'li', '2'li', 'kolye ve küpe', 'set', 'takım'.
    DİKKAT: motif içindeki 've' (ay_ve_yildiz) set DEĞİLDİR."""
    s = name.lower()
    # Türkçe ses uyumu: 2'li, 3'lü, 4'lü, 5'li  -> \d+_l[ıiuü]
    return bool(re.search(r"\d+_l[ıiuü]|kolye_ve_kupe|\bset\b|\btakim\b|takım", s))

def detect_stil(name):
    """Ürünün adında geçen TÜM stil ifadelerini toplar (ilkini değil)."""
    s = name.lower()
    out = []
    for k, v in [("uzun_y", "Uzun"), ("kisa_kuyruklu", "Kısa Kuyruklu"), ("tek_tas", "Tek Taş"),
                 ("ok_kupe", "Ok Uçlu"), ("ince", "İnce"), ("kalin", "Kalın"),
                 ("halka", "Halka"), ("klipsli", "Klipsli"), ("sallantili", "Sallantılı"),
                 ("y_kolye", "Y Formlu"), ("y_kupe", "Y Formlu"), ("u_kupe", "U Formlu")]:
        if k in s:
            out.append(v)
    return " ".join(dict.fromkeys(out)) or None

# ------------------------------------------------------------------ ana akış
def fix_nitelik(m):
    """Taş cinsini SADECE dosya adında açıkça yazıyorsa doldur; uydurma yok."""
    low = ascii_low(m["isim_kaynagi"])
    for a, b in (("kapmalali", "kaplamali"), ("kpe", "kupe"), ("renli", "renkli"),
                 ("renki", "renkli"), ("figurl", "figurlu"), ("yusufcu", "yusufcuk"),
                 ("temeali", "temali")):
        low = low.replace(a, b)          # yalnızca ÇÖZÜMLEME metni; ürün adı değişmez
    n = m["nitelik"]
    # motifler: öncelik sırasına göre (özgül motif önce, genel tema sonra)
    n["motifler"] = [k for k in MOTIF_PRIORITY if k in low]
    if "zirkon" in low:   c = "Zirkon"
    elif "inci" in low:   c = "İnci"
    elif "boncuk" in low: c = "Boncuk"
    elif "kristal" in low: c = "Kristal"
    elif "dogal" in low:  c = "Doğal Taş"
    else:                 c = None            # bilinmiyor -> söylenmez
    n["tas_cinsi_beyan"] = c                  # iddia edilebilir taş
    n["tas_cinsi"] = c if c else "Yok"        # enum alanı; geçersizse boş bırakılır
    n["tasli"] = ("tasli" in low) or ("tasi" in low) or bool(c)
    n["stone_phrase"] = STONE[c][0] if c in STONE else ("Taşlı" if n["tasli"] else None)
    n["mineli"] = "mineli" in low
    # --- Web Color: filtrelemede işe yarayan, ürünün baskın rengi
    col = n["renkler"][0] if n["renkler"] else None
    if col == "Zümrüt Yeşili": col = "Yeşil"
    # DİKKAT: 'gumus_renkli' = gümüş RENKLİ (metal rengi), çok renkli DEĞİL
    low_nc = low.replace("gumus_renkli", "").replace("gumus_renk_", "").replace("gumus_renk", "")
    multicolor = "renkli" in low_nc
    if multicolor:                                        wc = "Çok Renkli"
    elif "mineli" in low and col:                         wc = col
    elif n["metal"] == "Altın Kaplama":                   wc = "Altın"
    elif n["metal"] == "Gümüş":                           wc = "Gümüş"
    elif col:                                             wc = col
    else:                                                 wc = "Metalik"
    n["web_color"] = wc
    n["tas_rengi"] = col if (col and "mineli" not in low) else None


model = json.load(open(os.path.join(TMP, "glanssie_model.json"), encoding="utf-8"))
mapping = json.load(open(MAP, encoding="utf-8"))     # tek doğruluk kaynağı: SKU -> göreli yollar

# BDHA2025-230 tek bir ürün DEĞİL: görselleri ayrı ürünlere bölündü, o satır kaldırılır
model = [m for m in model if m["sku"] != "BDHA2025-230"]

for m in model:
    # görsel URL'lerini CDN sabitinden yeniden üret (model JSON'daki eski host'u kullanma)
    m["cdn_gorseller"] = [CDN + p for p in mapping[m["sku"]]]
    fix_nitelik(m)
    m["set_mi"] = is_set(m["isim_kaynagi"])
    m["nitelik"]["stil"] = detect_stil(m["isim_kaynagi"])
    m["title"] = compose_title(m)
    m["desc"] = compose_desc(m)

# --- Elle hazırlanan ek ürünler (İtalyan charm bileklikler, kod 245-261).
# Bunlar dosya adından türetilemez: her biri görselden tek tek incelenip isimlendirildi.
EK_YOL = os.path.join(os.path.dirname(MAP), "additional_products.json")
if os.path.exists(EK_YOL):
    for e in json.load(open(EK_YOL, encoding="utf-8")):
        e["cdn_gorseller"] = [CDN + p for p in mapping[e["sku"]]]
        e["set_mi"] = False
        e["stok"] = 100          # 230 satırının değerleri (kullanıcı onayı)
        e["maliyet"] = 65
        e["kategori_kodu"] = "2845"
        e["isim_kaynagi"] = e["title"]
        model.append(e)

# enum doğrulama
wb = openpyxl.load_workbook(TPL)
enums = {}
for cat, sh in ENUM_SHEET.items():
    ws = wb[sh]
    hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    d = {}
    for ci, h in enumerate(hdr, 1):
        if not h: continue
        vals = set()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, ci).value
            if v is not None and str(v).strip(): vals.add(str(v).strip())
        d[str(h).strip()] = vals
    enums[cat] = d

def attr_values(m):
    n = m["nitelik"]
    a = {}
    # --- ZORUNLU: barkod (klasör adı = barkod = stok kodu)
    a["Barkod"] = m["sku"]
    a["Marka"] = BRAND
    a["Kategori"] = m["kategori_kodu"]
    a["Para Birimi"] = "TRY"
    a["KDV Oranı"] = KDV
    a["Desi"] = DESI
    a["Sevkiyat Süresi"] = LEAD
    a["Stok Kodu"] = m["sku"]
    a["Model Kodu"] = m["sku"]
    a["Ürün Adı"] = m["title"]
    a["Ürün Açıklaması"] = m["desc"]
    a["Ürün Stok Adedi"] = int(m["stok"])
    a["Materyal"] = "Altın Kaplama" if n["metal"] == "Altın Kaplama" else "Çelik"
    a["Taş Cinsi"] = n["tas_cinsi"]
    a["Cinsiyet"] = "Kadın / Kız"
    a["Yaş Grubu"] = "Yetişkin"
    a["Garanti Süresi"] = "Garantisiz"
    a["Bakım Talimatları (Genel)"] = CARE
    if MENGEI: a["Menşei"] = MENGEI
    a["Web Color"] = n["web_color"]
    a["Renk"] = n["web_color"]
    # Küpe kategorisinde 'Model' alanı var: yalnızca dosya adından net çıkanlar yazılır
    if m["kategori"] == "Küpe":
        low = ascii_low(m["isim_kaynagi"])
        km = ("Halka" if "halka" in low else "Klipsli" if "klipsli" in low
              else "Sallantılı" if "sallantili" in low else None)
        if km: a["Model"] = km
    return a

# yazma
viol = []
sheet_headers = {}
for cat, sh in SHEET_OF_CAT.items():
    ws = wb[sh]
    sheet_headers[cat] = {str(ws.cell(1, c).value).strip(): c
                          for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}

row_ptr = collections.defaultdict(lambda: 3)
IMG_RX = re.compile(r"^Görsel (\d)$")
for m in model:
    cat = m["kategori"]; ws = wb[SHEET_OF_CAT[cat]]; hd = sheet_headers[cat]
    r = row_ptr[cat]; row_ptr[cat] += 1
    av = attr_values(m)
    for h, v in av.items():
        if h in hd:
            # enum doğrulaması: geçersiz değer YAZILMAZ (boş bırakılır), uydurma yok
            if h in enums[cat] and str(v) not in enums[cat][h]:
                viol.append((m["sku"], h, v))
                continue
            ws.cell(r, hd[h]).value = v
    # görseller
    for i, url in enumerate(m["cdn_gorseller"][:8], 1):
        if f"Görsel {i}" in hd: ws.cell(r, hd[f"Görsel {i}"]).value = url

wb.save(OUT)
print("Yazıldı:", OUT)
print("Satır dağılımı:", {c: row_ptr[c]-3 for c in SHEET_OF_CAT})
print("Toplam:", sum(row_ptr[c]-3 for c in SHEET_OF_CAT))
print("Enum ihlali:", len(viol))
for v in viol[:20]: print("   ", v)
json.dump(model, open(os.path.join(TMP, "glanssie_final.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
