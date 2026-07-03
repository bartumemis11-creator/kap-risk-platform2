# -*- coding: utf-8 -*-
"""
KAP Risk Tarama Uygulamasi
==========================
KAP (Kamuyu Aydinlatma Platformu, kap.org.tr) uzerinden secilen sirketlerin
bildirimlerini ceker ve risk iceren kayitlari tespit eder:

  - Icra / haciz / iflas / konkordato takipleri
  - Temerrut bildirimleri
  - Kredi derecelendirme (rating) notunun kotulesmesi
  - Yakin Izleme Pazari'na alinma / pazar degisikligi
  - SPK islem yasagi, brut takas, tedbir kararlari
  - Onemli davalar, idari para cezalari, sermaye kaybi (TTK 376) vb.

Sadece standart kutuphane kullanir (pip gerektirmez).
Cikti: konsol raporu + kap_risk_raporu.csv

Kullanim:  python kap_risk_tarama.py
"""

import csv
import gzip
import json
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

# ---------------------------------------------------------------- ayarlar ---

BASE = "https://www.kap.org.tr"

# Taranacak 10 firma (BIST hisse kodu ile)
TARGET_TICKERS = [
    "KONTR",   # Kontrolmatik Teknoloji
    "MARTI",   # Marti Otel
    "GSRAY",   # Galatasaray Sportif
    "FENER",   # Fenerbahce Futbol
    "VESTL",   # Vestel Elektronik
    "SASA",    # SASA Polyester
    "IHLAS",   # Ihlas Holding
    "THYAO",   # Turk Hava Yollari
    "TUPRS",   # Tupras
    "YKBNK",   # Yapi Kredi Bankasi
]

# Hangi yillarin bildirimleri cekilsin (None = guncel pencere)
YEARS = [None, 2026, 2025, 2024]

REQUEST_DELAY = 0.8   # istekler arasi bekleme (sn) - siteyi yormamak icin
TIMEOUT = 60
MAX_SANE_RESULT = 800  # bundan fazla kayit donerse filtre calismamis demektir

# ------------------------------------------------------- risk siniflamasi ---
# Her kural: (kategori, siddet, [anahtar kelimeler])  -- kelimeler normalize
# edilmis (kucuk harf, turkce karakterler sadelestirilmis) metinde aranir.

RISK_RULES = [
    ("İcra/İflas/Takip", "YÜKSEK",
     ["icra takib", "icra takip", "icra dairesi", "haciz", "iflas",
      "konkordato", "aleyhine takip", "kanuni takip", "yasal takip",
      "takip baslat"]),
    ("Temerrüt", "YÜKSEK",
     ["temerrut", "temerrud", "odeme guclugu", "odenmeme", "vadesinde odenme"]),
    ("Finansal Zorluk", "YÜKSEK",
     ["sermaye kaybi", "ttk 376", "376. madde", "borclanma araci odemesinde",
      "yeniden yapilandirma", "borc erteleme", "protesto", "karsiliksiz cek",
      "finansal guclugu"]),
    ("Pazar/İşlem Riski", "YÜKSEK",
     ["yakin izleme", "islem sirasi durdur", "piyasa oncesi islem",
      "sirasi kapatil", "borsa kotundan cikar", "kottan cikar"]),
    ("Pazar/İşlem Riski", "ORTA",
     ["islem yasagi", "pazar degisikligi", "brut takas", "tedbir",
      "kredili islem yasagi", "aciga satis yasagi", "tek fiyat"]),
    ("SPK/İdari", "ORTA",
     ["idari para cezasi", "suc duyurusu", "sorusturma", "spk denetim",
      "inceleme baslat", "kayyum"]),
    ("Dava/Hukuki", "ORTA",
     ["dava acil", "dava acti", "aleyhine dava", "davanin", "dava sureci",
      "tahkim", "arabuluculuk"]),
]

# Derecelendirme bildirimlerinde yon tespiti
RATING_KEYWORDS = ["derecelendirme", "kredi notu", "rating"]
RATING_NEGATIVE = ["dusur", "indirdi", "indirim", "negatif", "asagi",
                   "geri cek", "izlemeye al", "kotules", "temerrut", "d not",
                   "not cekil"]
RATING_POSITIVE = ["yukselt", "teyit", "pozitif", "duragan", "stabil"]

# Gurultu: risk olmayan rutin bildirimler (eslesirse atla)
NOISE_PATTERNS = ["devre kesici", "endeks sirketlerinde degisiklik",
                  "endeks degisikligi"]

# ------------------------------------------------------------- yardimcilar ---

TR_MAP = str.maketrans("çÇğĞıİöÖşŞüÜ", "ccggiioossuu")


def norm(s):
    """Turkce metni karsilastirma icin sadelestir."""
    if not s:
        return ""
    return s.translate(TR_MAP).lower()


def http_get(url, retries=3):
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/126.0.0.0 Safari/537.36"),
        "Accept": "*/*",
        "Accept-Encoding": "gzip",
        "Accept-Language": "tr",
    }
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=headers)
            resp = urllib.request.urlopen(req, timeout=TIMEOUT)
            body = resp.read()
            if resp.headers.get("Content-Encoding") == "gzip":
                body = gzip.decompress(body)
            return body.decode("utf-8", "ignore")
        except Exception as exc:  # aglar kaprisli olabilir -> tekrar dene
            last_err = exc
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Istek basarisiz: {url} -> {last_err}")


def extract_flight_array(html, anchor='\\"data\\":['):
    """Next.js SSR sayfasina gomulu (escape edilmis) JSON dizisini cikar."""
    start = html.find(anchor)
    if start < 0:
        return []
    region = html[start:start + 3_000_000]
    txt = (region.replace('\\\\"', "@@Q@@")
                 .replace('\\"', '"')
                 .replace("@@Q@@", '\\"'))
    a = txt.find("[")
    depth, in_str, esc = 0, False, False
    for i, ch in enumerate(txt[a:], a):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(txt[a:i + 1])
                except json.JSONDecodeError:
                    return []
    return []


def parse_date(s):
    try:
        return datetime.strptime(s, "%d.%m.%Y %H:%M:%S")
    except Exception:
        return datetime.min

# ------------------------------------------------------------ veri cekimi ---


def fetch_member_directory():
    """KAP uye rehberini (sirket adi / hisse kodu / mkkMemberOid) cek."""
    print("KAP sirket listesi indiriliyor...")
    html = http_get(f"{BASE}/tr/bildirim-sorgu")
    pattern = re.compile(
        r'\\"mkkMemberOid\\":\\"([0-9a-fA-F]+)\\",'
        r'\\"kapMemberTitle\\":\\"(.*?)\\",'
        r'.*?\\"stockCode\\":\\"(.*?)\\"')
    members = {}
    for oid, title, stock in pattern.findall(html):
        try:  # \uXXXX kacislarini coz, duz UTF-8 metni bozma
            title = json.loads('"' + title.replace('"', '\\"') + '"')
        except json.JSONDecodeError:
            pass
        for code in stock.replace(";", ",").split(","):
            code = code.strip().upper()
            if code and code != "-":
                members[code] = {"oid": oid, "title": title, "stock": code}
    print(f"  {len(members)} hisse kodu cozumlendi.")
    return members


def fetch_company_disclosures(member):
    """Bir sirketin bildirimlerini yil yil cekip birlestir."""
    seen = {}
    for yr in YEARS:
        url = (f"{BASE}/tr/bildirim-sorgu-sonuc?srcbar=Y&cmp=Y&cat=2"
               f"&m={member['oid']}&t=X&slf=ALL")
        if yr:
            url += f"&yr={yr}"
        try:
            html = http_get(url)
        except RuntimeError as exc:
            print(f"    uyari: {exc}")
            continue
        items = extract_flight_array(html)
        if len(items) > MAX_SANE_RESULT:
            # filtre uygulanmamis genel dokum donmus -> guvenme, atla
            continue
        for it in items:
            basic = it.get("disclosureBasic") or it
            idx = basic.get("disclosureIndex")
            if idx and idx not in seen:
                seen[idx] = basic
        time.sleep(REQUEST_DELAY)
    return sorted(seen.values(),
                  key=lambda b: parse_date(b.get("publishDate", "")),
                  reverse=True)

# ------------------------------------------------------- risk tespiti -------


def classify_risk(basic):
    """Bildirimi risk kurallarina gore sinifla. None = risk degil."""
    text = norm(" | ".join(filter(None, [
        basic.get("title"), basic.get("summary"), basic.get("subject"),
        basic.get("ruleTypeTerm"),
    ])))
    if not text:
        return None
    for noise in NOISE_PATTERNS:
        if noise in text:
            return None

    # 1) Derecelendirme bildirimleri: yalnizca kotulesme/belirsizlik riskli
    if any(k in text for k in RATING_KEYWORDS):
        if any(k in text for k in RATING_NEGATIVE):
            return ("Derecelendirme Notu Kötüleşmesi", "YÜKSEK",
                    "not dusurme/negatif gorunum ifadesi tespit edildi")
        if any(k in text for k in RATING_POSITIVE):
            return None  # teyit/yukseltme -> risk degil
        return ("Derecelendirme (yön belirsiz)", "ORTA",
                "derecelendirme bildirimi - ozetten yon anlasilamadi, "
                "detay kontrol edilmeli")

    # 2) Genel kurallar
    for category, severity, keywords in RISK_RULES:
        for kw in keywords:
            if kw in text:
                return (category, severity, f"anahtar kelime: '{kw}'")
    return None

# ------------------------------------------------------------------ rapor ---


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    print("=" * 78)
    print("KAP RİSK TARAMA — takip / temerrüt / derecelendirme / pazar riski")
    print("=" * 78)

    directory = fetch_member_directory()

    companies = []
    for ticker in TARGET_TICKERS:
        m = directory.get(ticker)
        if m:
            companies.append(m)
        else:
            print(f"  UYARI: {ticker} kodu KAP listesinde bulunamadi, atlaniyor.")

    all_findings = []
    for i, member in enumerate(companies, 1):
        print(f"\n[{i}/{len(companies)}] {member['stock']} — "
              f"{member['title'][:55]}")
        disclosures = fetch_company_disclosures(member)
        print(f"    {len(disclosures)} bildirim tarandi", end="")
        findings = []
        for basic in disclosures:
            hit = classify_risk(basic)
            if hit:
                category, severity, reason = hit
                findings.append({
                    "hisse": member["stock"],
                    "sirket": member["title"],
                    "tarih": basic.get("publishDate", ""),
                    "kategori": category,
                    "siddet": severity,
                    "baslik": basic.get("title") or "",
                    "ozet": basic.get("summary") or "",
                    "gerekce": reason,
                    "link": f"{BASE}/tr/Bildirim/{basic.get('disclosureIndex')}",
                })
        all_findings.extend(findings)
        print(f" -> {len(findings)} riskli kayit")

    # ---- konsol raporu ----
    print("\n" + "=" * 78)
    print("SONUÇ RAPORU")
    print("=" * 78)
    order = {"YÜKSEK": 0, "ORTA": 1}
    by_company = {}
    for f in all_findings:
        by_company.setdefault(f["hisse"], []).append(f)

    for member in companies:
        code = member["stock"]
        findings = sorted(by_company.get(code, []),
                          key=lambda f: (order.get(f["siddet"], 9),
                                         -parse_date(f["tarih"]).timestamp()
                                         if f["tarih"] else 0))
        print(f"\n{'-' * 78}")
        print(f"{code} — {member['title']}")
        if not findings:
            print("   Riskli kayit tespit edilmedi.")
            continue
        yuksek = sum(1 for f in findings if f["siddet"] == "YÜKSEK")
        print(f"   {len(findings)} riskli kayit ({yuksek} YÜKSEK)")
        for f in findings:
            print(f"   [{f['siddet']:^6}] {f['tarih'][:10]}  {f['kategori']}")
            detail = f["ozet"] or f["baslik"]
            if detail:
                print(f"            {detail[:95]}")
            print(f"            {f['gerekce']}  |  {f['link']}")

    # ---- ozet tablo ----
    print(f"\n{'=' * 78}")
    print(f"{'HİSSE':<8}{'YÜKSEK':>8}{'ORTA':>8}{'TOPLAM':>8}   DURUM")
    print("-" * 78)
    for member in companies:
        fs = by_company.get(member["stock"], [])
        y = sum(1 for f in fs if f["siddet"] == "YÜKSEK")
        o = len(fs) - y
        durum = ("*** RİSKLİ ***" if y else ("izlenmeli" if o else "temiz"))
        print(f"{member['stock']:<8}{y:>8}{o:>8}{len(fs):>8}   {durum}")

    # ---- csv ----
    out = "kap_risk_raporu.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=[
            "hisse", "sirket", "tarih", "siddet", "kategori",
            "baslik", "ozet", "gerekce", "link"])
        writer.writeheader()
        for f in sorted(all_findings,
                        key=lambda f: (f["hisse"], order.get(f["siddet"], 9))):
            writer.writerow(f)
    print(f"\nDetayli rapor kaydedildi: {out}")


if __name__ == "__main__":
    main()
