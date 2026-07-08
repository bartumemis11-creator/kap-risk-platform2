# -*- coding: utf-8 -*-
"""kap_rapor_mail birim testleri — ağ / SMTP gerektirmez (saf yardımcılar)."""
import os
import sys

import kap_rapor_mail as mail
import kap_risk_app as app


def _sahte_sonuc(hisse, not_, skor, seviye, bulgular):
    return {
        "member": {"hisse": hisse, "unvan": f"{hisse} A.Ş.",
                   "oid": "1", "islem": "1"},
        "taranan": 42, "findings": bulgular, "improvements": [],
        "skor": skor, "not": not_, "seviye": seviye,
        "emoji": "🔴", "renk": "#000", "veri_hatasi": 0,
    }


def _sahte_bulgu(kategori, baslik):
    return {"agirlik": 10, "guncellik": 1.0, "kategori": kategori,
            "baslik": baslik, "tarih_str": "2026-07-01", "siddet": "KRİTİK"}


def test_izleme_listesi_env_oncelikli():
    os.environ["IZLEME_LISTESI"] = "aselS, thyao,,tuprS"
    try:
        assert mail.izleme_listesi() == ["ASELS", "THYAO", "TUPRS"]
    finally:
        del os.environ["IZLEME_LISTESI"]


def test_izleme_listesi_bos_ise_varsayilan():
    os.environ.pop("IZLEME_LISTESI", None)
    # izleme_listesi.txt repoda mevcut; en az bir kod dönmeli
    assert len(mail.izleme_listesi()) >= 1


def test_csv_bytes_bom_ve_baslik():
    veri = mail.csv_bytes([{k: "x" for k in mail.CSV_ALANLARI}])
    assert veri.startswith(b"\xef\xbb\xbf")           # UTF-8 BOM
    metin = veri.decode("utf-8-sig")
    assert metin.splitlines()[0] == ",".join(mail.CSV_ALANLARI)
    assert len(metin.splitlines()) == 2               # başlık + 1 satır


def test_md_kalin_html_donusumu():
    assert mail._md_satir_html("a **b** c") == "a <strong>b</strong> c"


def test_ozet_html_paragraf_uretir():
    html = mail.ozet_html("ilk **satır**\n\nikinci satır")
    assert html.count("<p") == 2
    assert "<strong>satır</strong>" in html


def test_mail_govdesi_konu_ve_kritik():
    results = [
        _sahte_sonuc("AAA", "E", 88, "ÇOK YÜKSEK",
                     [_sahte_bulgu("Temerrüt", "Ödeme gerçekleşmedi")]),
        _sahte_sonuc("BBB", "A", 0, "TEMİZ", []),
    ]
    konu, duz, html = mail.mail_govdesi(results, (2024, 2025, 2026), True)
    assert "KAP Risk Raporu" in konu
    assert "AAA" in konu and "acil" in konu           # kritik şirket konuda
    assert "**" not in duz                            # düz metinde md kalmamış
    assert "<table" in html and "AAA" in html


def test_mail_olustur_iki_ek_ve_alicilar():
    msg = mail.mail_olustur(
        "konu", "duz", "<p>html</p>", "gonderen@x.com",
        ["a@k.com", "b@k.com"], b"excel", b"csv", "20260707")
    assert msg["Subject"] == "konu"
    assert msg["To"] == "a@k.com, b@k.com"
    ekler = [p for p in msg.iter_attachments()]
    assert len(ekler) == 2
    adlar = sorted(p.get_filename() for p in ekler)
    assert adlar == ["KAP_Risk_Bulgular_20260707.csv",
                     "KAP_Risk_Raporu_20260707.xlsx"]


def test_email_regex():
    assert app.EMAIL_RE.match("ad.soyad@ornek.com")
    assert not app.EMAIL_RE.match("eksik@alan")
    assert not app.EMAIL_RE.match("bosluk lu@x.com")


def test_talep_gecersiz_email_reddedilir():
    ok, msg = app.talep_mail_gonder("bozuk-adres", [{"findings": []}],
                                    (2026,), True, b"xls")
    assert ok is False and "Geçerli" in msg


def test_talep_bos_sonuc_reddedilir():
    # geçerli adres ama tarama yok → SMTP'ye hiç dokunmadan reddetmeli
    ok, msg = app.talep_mail_gonder("a@b.com", [], (2026,), True, b"xls")
    assert ok is False and "tarama" in msg.lower()


if __name__ == "__main__":
    fails = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"OK   {name}")
            except AssertionError as exc:
                fails += 1
                print(f"FAIL {name}: {exc}")
    print(f"\n{fails} başarısız")
    sys.exit(1 if fails else 0)
