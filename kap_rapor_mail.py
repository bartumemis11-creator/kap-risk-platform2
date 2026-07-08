# -*- coding: utf-8 -*-
"""
KAP Risk Raporu — günlük e-posta gönderici
==========================================
Seçili izleme listesindeki BIST şirketlerini tarar, biçimlendirilmiş Excel
raporu + CSV üretir ve bunları e-posta ekinde (yönetici özeti gövdede) gönderir.
Tarama/skorlama/rapor mantığı TEK kaynaktan — `kap_risk_app.py` çekirdeğinden —
gelir; böylece web arayüzü, CLI ve mail çıktısı asla sessizce ayrışmaz.

Zamanlama GitHub Actions'tadır (bkz. `.github/workflows/gunluk-rapor.yml`);
bu betik yalnızca "tara → rapor üret → gönder" akışını sağlar.

Varsayılan akış: umut.okan1@gmail.com (Gmail) → umutbasbay@hotmail.com.
Adresler ve sunucu koda gömülü varsayılanlardır; ortam değişkenleriyle
değiştirilebilir. Parola koda ASLA girmez.

Ortam değişkenleri (GitHub Secrets / Variables)
-----------------------------------------------
Zorunlu:
  SMTP_PASS   gönderen hesabın parolası — Gmail için "uygulama parolası"
İsteğe bağlı (varsayılanları geçersiz kılar):
  SMTP_HOST   e-posta sunucusu (varsayılan: smtp.gmail.com)
  SMTP_PORT   varsayılan 587 (STARTTLS)
  SMTP_USER   gönderen kullanıcı adı; boşsa MAIL_FROM kullanılır
  MAIL_FROM   gönderen adresi (varsayılan: umut.okan1@gmail.com)
  MAIL_TO     alıcı(lar), virgülle ayrılmış (varsayılan: umutbasbay@hotmail.com)
  SMTP_OAUTH_TOKEN  verilirse XOAUTH2 ile kimlik doğrulanır (parola yerine)
  IZLEME_LISTESI    virgülle ayrılmış hisse kodları; boşsa `izleme_listesi.txt`,
                    o da yoksa çekirdeğin DEFAULT_TICKERS listesi kullanılır
  TARAMA_YIL        tarama başlangıç yılı; boşsa (bugün - 2)
  DERIN             "0"/"false" değilse derin mod (varsayılan: derin)
"""
import base64
import csv
import io
import os
import smtplib
import ssl
import sys
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kap_risk_app as core  # noqa: E402  (kurallar/veri/rapor katmanı çekirdekten)

CSV_ALANLARI = ("hisse", "sirket", "tarih_str", "siddet",
                "kategori", "baslik", "ozet", "gerekce", "link")

# Varsayılan gönderim yapılandırması — ortam değişkenleri (GitHub Secrets)
# tanımlıysa onlar öncelik kazanır. Parola ASLA burada tutulmaz; yalnızca
# SMTP_PASS gizli değişkeniyle verilir (Gmail için "uygulama parolası").
VARSAYILAN_SMTP_HOST = "smtp.gmail.com"
VARSAYILAN_FROM = "umut.okan1@gmail.com"     # gönderen (Gmail)
VARSAYILAN_TO = "umutbasbay@hotmail.com"      # alıcı (Hotmail)


# ───────────────────────────────────────────────────── yapılandırma ──

def _env(ad: str, varsayilan: str = "") -> str:
    return (os.environ.get(ad) or varsayilan).strip()


def _dogru_mu(deger: str) -> bool:
    return deger.strip().lower() not in ("", "0", "false", "hayir", "hayır", "no")


def izleme_listesi() -> list:
    """İzlenecek hisse kodları: ortam değişkeni → dosya → çekirdek varsayılanı."""
    ham = _env("IZLEME_LISTESI")
    if not ham:
        yol = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "izleme_listesi.txt")
        if os.path.exists(yol):
            with open(yol, encoding="utf-8") as fh:
                # satır bazlı; '#' ile başlayan satırlar yorum
                ham = ",".join(s.split("#", 1)[0].strip() for s in fh)
    kodlar = [h.strip().upper() for h in ham.replace("\n", ",").split(",")
              if h.strip()]
    return kodlar or list(core.DEFAULT_TICKERS)


# ─────────────────────────────────────────────────────────── tarama ──

def taramayi_calistir(kodlar: list, baslangic_yili: int, derin: bool):
    """İzleme listesini tarar → (results, years, csv_satirlari)."""
    print("KAP üye rehberi indiriliyor...", flush=True)
    directory = core.fetch_member_directory()
    istenen = set(kodlar)

    members, bulunan = [], set()
    for row in directory.itertuples():
        row_kodlar = set(str(row.kodlar).split(","))
        kesisim = row_kodlar & istenen
        if kesisim:
            members.append({"hisse": row.hisse, "unvan": row.unvan,
                            "oid": row.oid, "islem": row.islem})
            bulunan |= kesisim
    eksik = sorted(istenen - bulunan)
    if eksik:
        print("  UYARI - rehberde bulunamadı:", ", ".join(eksik), flush=True)

    years = tuple(range(baslangic_yili, date.today().year + 1))
    results, csv_satirlari = [], []
    for i, m in enumerate(members, 1):
        print(f"[{i}/{len(members)}] {m['hisse']} — {m['unvan'][:55]}", flush=True)
        r = core.scan_company(m, years, deep=derin)
        results.append(r)
        print(f"    {r['taranan']} bildirim → {len(r['findings'])} bulgu | "
              f"skor {r['skor']:.0f}/100 ({r['not']}-{r['seviye']})", flush=True)
        for f in sorted(r["findings"],
                        key=lambda f: core.SEVERITY_ORDER.get(f["siddet"], 9)):
            csv_satirlari.append({k: f.get(k, "") for k in CSV_ALANLARI})
    return results, years, csv_satirlari


def csv_bytes(satirlar: list) -> bytes:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=list(CSV_ALANLARI))
    w.writeheader()
    w.writerows(satirlar)
    # Excel'in Türkçe karakterleri doğru açması için UTF-8 BOM
    return b"\xef\xbb\xbf" + buf.getvalue().encode("utf-8")


# ──────────────────────────────────────────────────── e-posta gövdesi ──

def _md_satir_html(satir: str) -> str:
    """Yönetici özetindeki **kalın** işaretlerini <strong>'a çevirir.

    Metin zaten &nbsp; gibi HTML varlıkları içerdiğinden kaçışlamıyoruz;
    içerik güvenilir (çekirdek üretiyor, kullanıcı girdisi değil)."""
    parcalar = satir.split("**")
    html = ""
    for i, p in enumerate(parcalar):
        html += p if i % 2 == 0 else f"<strong>{p}</strong>"
    return html


def ozet_html(ozet_md: str) -> str:
    paragraflar = [f"<p style='margin:0 0 10px'>{_md_satir_html(s)}</p>"
                   for s in ozet_md.split("\n\n") if s.strip()]
    return "\n".join(paragraflar)


def risk_tablosu_html(results: list) -> str:
    """En riskliden temize sıralı özet tablo (okunabilir HTML)."""
    sirali = sorted(results, key=lambda r: -r["skor"])
    satirlar = []
    for r in sirali:
        m = r["member"]
        en_agir = max(r["findings"],
                      key=lambda f: f["agirlik"] * f["guncellik"], default=None)
        gerekce = ""
        if en_agir:
            gerekce = f"{en_agir['kategori']}: {en_agir['baslik'][:70]}"
        satirlar.append(
            "<tr>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>"
            f"<b>{m['hisse']}</b></td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
            f"text-align:center'>{r['emoji']} {r['not']}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
            f"text-align:right'>{r['skor']:.0f}/100</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
            f"text-align:center'>{len(r['findings'])}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee;"
            f"color:#555'>{gerekce}</td>"
            "</tr>")
    return (
        "<table style='border-collapse:collapse;width:100%;font-size:13px;"
        "font-family:Segoe UI,Arial,sans-serif'>"
        "<thead><tr style='background:#1F3864;color:#fff'>"
        "<th style='padding:8px 10px;text-align:left'>Hisse</th>"
        "<th style='padding:8px 10px'>Not</th>"
        "<th style='padding:8px 10px;text-align:right'>Skor</th>"
        "<th style='padding:8px 10px'>Bulgu</th>"
        "<th style='padding:8px 10px;text-align:left'>En ağır sinyal</th>"
        "</tr></thead><tbody>"
        + "".join(satirlar) + "</tbody></table>")


def mail_govdesi(results: list, years: tuple, derin: bool) -> tuple:
    """(konu, düz_metin, html) döndürür."""
    tarih = datetime.now().strftime("%d.%m.%Y")
    kritik = [r for r in results if r["not"] in ("D", "E")]
    toplam_bulgu = sum(len(r["findings"]) for r in results)
    mod = "derin" if derin else "hızlı"

    konu = f"KAP Risk Raporu — {tarih}"
    if kritik:
        adlar = ", ".join(r["member"]["hisse"]
                          for r in sorted(kritik, key=lambda r: -r["skor"]))
        konu += f" · 🔴 {len(kritik)} acil ({adlar})"
    else:
        konu += f" · {toplam_bulgu} sinyal"

    ozet_md = core.executive_summary_text(results)
    baslik = (f"<h2 style='color:#1F3864;font-family:Segoe UI,Arial,sans-serif;"
              f"margin:0 0 4px'>🛡️ KAP Risk İzleme Raporu</h2>"
              f"<p style='color:#777;font-family:Segoe UI,Arial,sans-serif;"
              f"margin:0 0 16px;font-size:13px'>{tarih} · {years[0]}–{years[-1]} "
              f"dönemi · {mod} analiz · {len(results)} şirket</p>")
    not_metni = ("<p style='color:#999;font-size:12px;margin-top:18px;"
                 "font-family:Segoe UI,Arial,sans-serif'>Ayrıntılı 5+ sayfalık "
                 "rapor ekteki Excel dosyasındadır; tüm bulgular ekteki CSV'de. "
                 "Bu araç yatırım tavsiyesi değildir; karar öncesi bildirim "
                 "asıllarını KAP'ta doğrulayınız.</p>")
    html = (f"<div style='max-width:820px'>{baslik}{ozet_html(ozet_md)}"
            f"<h3 style='color:#1F3864;font-family:Segoe UI,Arial,sans-serif;"
            f"margin:20px 0 8px'>Şirket Özeti</h3>"
            f"{risk_tablosu_html(results)}{not_metni}</div>")

    # düz metin alternatifi (** işaretleri temizlenmiş)
    duz = ozet_md.replace("**", "").replace("&nbsp;", " ")
    duz = (f"KAP Risk İzleme Raporu — {tarih}\n{years[0]}–{years[-1]} · "
           f"{mod} analiz · {len(results)} şirket\n\n{duz}\n\n"
           "Ayrıntılı rapor ekteki Excel'de; tüm bulgular ekteki CSV'de.\n"
           "Bu araç yatırım tavsiyesi değildir.")
    return konu, duz, html


# ──────────────────────────────────────────────────────── gönderim ──

def mail_olustur(konu, duz, html, gonderen, alicilar,
                 excel_bytes, csv_veri, tarih_str) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = konu
    msg["From"] = formataddr(("KAP Risk İzleme", gonderen))
    msg["To"] = ", ".join(alicilar)
    msg["Date"] = formatdate(localtime=True)
    msg.set_content(duz)
    msg.add_alternative(html, subtype="html")

    msg.add_attachment(
        excel_bytes, maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"KAP_Risk_Raporu_{tarih_str}.xlsx")
    msg.add_attachment(
        csv_veri, maintype="text", subtype="csv",
        filename=f"KAP_Risk_Bulgular_{tarih_str}.csv")
    return msg


def gonder(msg: EmailMessage):
    host = _env("SMTP_HOST", VARSAYILAN_SMTP_HOST)
    port = int(_env("SMTP_PORT", "587"))
    user = _env("SMTP_USER") or _env("MAIL_FROM") or VARSAYILAN_FROM
    parola = _env("SMTP_PASS")
    oauth = _env("SMTP_OAUTH_TOKEN")
    if not (parola or oauth):
        raise SystemExit("HATA: SMTP_PASS (Gmail uygulama parolası) veya "
                         "SMTP_OAUTH_TOKEN ortam değişkeni zorunlu.")

    baglam = ssl.create_default_context()
    with smtplib.SMTP(host, port, timeout=60) as s:
        s.ehlo()
        s.starttls(context=baglam)
        s.ehlo()
        if oauth:
            # XOAUTH2 — kurumsal olmayan Outlook'ta basic auth kapalıysa
            xoauth = f"user={user}\x01auth=Bearer {oauth}\x01\x01"
            s.docmd("AUTH", "XOAUTH2 "
                    + base64.b64encode(xoauth.encode()).decode())
        else:
            s.login(user, parola)
        s.send_message(msg)


# ─────────────────────────────────────────────────────────── main ──

def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    alicilar = [a.strip() for a in
                (_env("MAIL_TO") or VARSAYILAN_TO).replace("\n", ",").split(",")
                if a.strip()]
    if not alicilar:
        raise SystemExit("HATA: MAIL_TO (alıcı adresi) zorunlu.")
    gonderen = _env("MAIL_FROM") or _env("SMTP_USER") or VARSAYILAN_FROM

    kodlar = izleme_listesi()
    baslangic = int(_env("TARAMA_YIL") or (date.today().year - 2))
    derin = _dogru_mu(_env("DERIN", "1"))
    print(f"İzleme listesi ({len(kodlar)}): {', '.join(kodlar)}", flush=True)

    results, years, csv_satirlari = taramayi_calistir(kodlar, baslangic, derin)
    if not results:
        raise SystemExit("HATA: İzleme listesindeki hiçbir hisse KAP "
                         "rehberinde bulunamadı; mail gönderilmedi.")

    excel = core.build_excel(results, years, derin)
    csv_veri = csv_bytes(csv_satirlari)
    tarih_str = datetime.now().strftime("%Y%m%d")
    konu, duz, html = mail_govdesi(results, years, derin)

    msg = mail_olustur(konu, duz, html, gonderen, alicilar,
                       excel, csv_veri, tarih_str)
    gonder(msg)
    print(f"\n✓ Rapor gönderildi → {', '.join(alicilar)}", flush=True)
    print(f"  Konu: {konu}", flush=True)


if __name__ == "__main__":
    main()
