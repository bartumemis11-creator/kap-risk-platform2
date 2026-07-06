# Altin vaka regresyonu: oturum boyunca YAKALANMIS gercek sinyaller (ag yok)
import sys
sys.path.insert(0, r"C:\Users\MSİ\Desktop\ccdene1")
sys.stdout.reconfigure(encoding="utf-8")
import kap_risk_app as app

M = {"hisse": "T", "unvan": "T A.Ş.", "oid": "x"}
def b(t, s="", pub="01.06.2026 10:00:00"):
    return {"title": t, "summary": s, "companyTitle": "T A.Ş.",
            "publishDate": pub, "disclosureIndex": 1}

# (aciklama, basic, detay, beklenen kategori | "IYILESME" | None)
GOLDEN = [
 ("MKK gerceklesmeyen itfa", b("Merkezi Kayıt Kuruluşu A.Ş. Duyurusu",
   "Gerçekleşmeyen İtfa/Kupon/Getiri Ödemesi"), "", "temerrut"),
 ("gozalti pazari", b("Borsa İstanbul A.Ş. Duyurusu",
   "Borçlanma Araçlarının Gözaltı Pazarı'na Alınması"), "", "yakin_izleme"),
 ("yakin izleme alinma", b("Pazar Değişikliği",
   "Payların Yakın İzleme Pazarına Alınması"), "", "yakin_izleme"),
 ("dunya bankasi yaptirim", b("Özel Durum Açıklaması (Genel)",
   "Dünya Bankası Grubu Tarafından Uygulanan Yaptırım Hakkında"), "", "regulator"),
 ("nisab erteleme", b("Genel Kurul İşlemlerine İlişkin Bildirim",
   "Asgari Toplantı Nisabının Sağlanamaması Nedeni ile Ertelenen GK"), "", "yonetim"),
 ("tedbir uygulanmasi", b("Borsa İstanbul A.Ş. Duyurusu",
   "Yatırım Aracı Bazında Tedbir Uygulanması"), "", "piyasa_tedbir"),
 ("MARTI kredi yapilandirma", b("Özel Durum Açıklaması (Genel)",
   "Denizbank ve Deniz Fact. Kredi Yapılandırmaları Protokol Görüşmeleri Hk."), "", "yapilandirma"),
 ("SASA covenant", b("Özel Durum Açıklaması (Genel)",
   "PTA Yatırımı kredilerinin raporlanması hakkında"), "", "yapilandirma"),
 ("IHLAS rapor gecikmesi", b("Finansal Rapor",
   "31.12.2023 dönemine ait konsolide finansal rapor",
   "20.05.2024 22:00:00"), "", "denetim"),
 ("KONTR likidite (derin)", b("Özel Durum Açıklaması (Genel)", "Güncel Durum"),
   "nakit akışındaki bozulma ve likidite sıkışıklığı yaşanmaktadır", "temerrut"),
 ("VESTL FYY", b("Özel Durum Açıklaması (Genel)",
   "Finansal Yeniden Yapılandırma Başvurusu Hakkında"), "", "yapilandirma"),
 ("FENER UEFA limit", b("Özel Durum Açıklaması (Genel)",
   "UEFA Sürdürülebilirlik Talimatı Kapsamında Limit Aşımı Hakkında"), "", "regulator"),
 ("GM yrd ayrilma", b("Özel Durum Açıklaması (Genel)",
   "Genel Müdür Yardımcısının ayrılması"), "", "yonetim"),
 ("YKBNK TGA satisi", b("Özel Durum Açıklaması (Genel)",
   "Tahsili Gecikmiş Alacak Portföyü Alımı-İhale"), "", "varlik_satisi"),
 ("VIOP vade acilmamasi", b("VİOP Diğer Duyurular",
   "Payına dayalı vadeli işlem sözleşmelerinde yeni vade aylarının işleme açılmaması"), "", "yakin_izleme"),
 ("rating dusurme (derin)", b("Kredi Derecelendirme Notu"),
   "kredi notunu BBB'den BB'ye düşürmüştür, görünüm negatif", "derecelendirme"),
 ("YIP cikis = iyilesme", b("Pazar Değişikliği",
   "Payların Yakın İzleme Pazarı'ndan çıkarılarak Ana Pazar'a alınması"), "", "IYILESME"),
 ("konkordato sona erdi = RISK", b("Özel Durum Açıklaması",
   "Konkordato mühleti sona erdi, iflas süreci başladı"), "", "iflas"),
 # ── piyasa taramasında bulunan sızıntılar (kapatıldı) ──
 ("varlik SATIMI varyanti", b("Maddi Duran Varlık Satımı",
   "Nurol Tower Ofis Satışı"), "", "varlik_satisi"),
 ("sirketin uyarilmasi", b("Şirketin Uyarılması",
   "Şirketin uyarılması"), "", "regulator"),
 ("ipotek tesisi", b("Özel Durum Açıklaması (Genel)",
   "Kredi Teminatı Kapsamında İpotek Tesis Edilmesi"), "", "varlik_satisi"),
 ("sube kapanisi", b("Özel Durum Açıklaması (Genel)",
   "Şube Kapanışı Hakkında"), "", "faaliyet"),
 ("rutin kupon = temiz", b("Pay Dışında Sermaye Piyasası Aracı İşlemleri",
   "TRF ISIN Kodlu Finansman Bonosunun 3. Kupon Ödemesi"), "", None),
 ("rutin GK daveti = temiz", b("Genel Kurul İşlemlerine İlişkin Bildirim",
   "2025 Yılı Olağan Genel Kurul Toplantısı Daveti"), "", None),
]
fail = 0
for name, basic, deep, exp in GOLDEN:
    r = app.classify(basic, M, deep)
    got = None if r is None else ("IYILESME" if r["iyilesme"] else r["kategori_id"])
    ok = got == exp
    fail += 0 if ok else 1
    print(("OK  " if ok else "FAIL"), f"{name:28s} beklenen={str(exp):14s}",
          "" if ok else f"bulunan={got}")
print(f"\n{len(GOLDEN)} altin vaka, {fail} basarisiz")
sys.exit(1 if fail else 0)
