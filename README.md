# 🛡️ KAP Risk İzleme ve Erken Uyarı Platformu

KAP (Kamuyu Aydınlatma Platformu, [kap.org.tr](https://www.kap.org.tr)) bildirimlerini tarayarak
BIST şirketlerinin kötüye gidişine işaret eden sinyalleri tespit eden, ağırlıklı risk skoru üreten
ve yapılandırılmış bir risk yönetimi raporu (Excel) sunan Streamlit uygulaması.

## Tespit edilen risk sinyalleri

| Kategori | Ağırlık |
|---|---|
| İflas / Tasfiye / Konkordato | 10 |
| Temerrüt / Ödeme Performansı (likidite sıkışıklığı dâhil) | 10 |
| Yakın İzleme Pazarı / Kotasyon Riski | 9 |
| İhaleye Fesat / Yolsuzluk / Kayyum / Adli süreçler | 9 |
| Finansal Yeniden Yapılandırma / Sermaye Kaybı (TTK 376) | 8 |
| Denetçi Görüşü / İşletme Sürekliliği Şüphesi | 8 |
| Kredi Derecelendirme Notu Kötüleşmesi (yön tespitli) | 7 |
| İcra / Haciz Takipleri | 7 |
| Regülatör Cezaları (SPK, BDDK, EPDK, Rekabet K., vergi) | 7 |
| Faaliyet / Üretim Riski (grev, üretim durdurma, lisans iptali) | 5 |
| Yönetim / Genel Kurul Riski (istifalar, nisab sağlanamaması) | 5 |
| Varlık Satışı / Nakit Yaratma (duran varlık, iştirak, TGA satışı) | 4 |
| Kârlılık / Temettü Sinyali (kâr dağıtılmaması, dönem zararı) | 4 |
| Önemli Dava / Tahkim | 4 |
| Piyasa Tedbirleri (brüt takas, VBTS, açığa satış yasağı) | 3 |

Borçlanma araçları tarafında MKK'nın **"Gerçekleşmeyen İtfa/Kupon/Getiri Ödemesi"**
duyuruları, borçlanma araçlarının **Gözaltı Pazarı'na alınması**, VİOP'ta **yeni vade
ayının işleme açılmaması** ve **temerrüt alışı/açığı** bildirimleri de yakalanır.

Ek olarak **Google News** üzerinden ulusal basındaki risk temalı haberler (temerrüt, konkordato,
kayyum, işlem yasağı…) bilgilendirme amaçlı taranır; risk skoruna dâhil edilmez.

## Özellikler

- 🏢 Şirket listesi KAP'tan otomatik çekilir (~1.000+ hisse), arama destekli çoklu seçim + "Tümünü seç"
- 📅 Elle tarih aralığı girme (2015'e kadar geriye dönük)
- 🔎 İki analiz derinliği: hızlı (başlık/özet) ve derin (bildirim tam metinleri)
- 🧮 Ağırlık × güncellik × sönümleme tabanlı 0-100 risk skoru, A–E notu
- 📰 KAP formatında riskli bildirim akışı — tıklayınca bildirim KAP'ta açılır
- 🆕 Bir önceki taramaya göre yeni bulguların işaretlenmesi
- ↗️ İyileşme sinyallerinin (tedbir kaldırma, pazar çıkışı, lehte karar) ayrı raporlanması
- 📥 5+ sayfalık biçimlendirilmiş Excel raporu ve CSV dışa aktarımı
- 🌐 Hız sınırlayıcı + devre kesici + tekrar deneme ile KAP dostu, dayanıklı veri çekimi

## Kurulum ve çalıştırma

```bash
pip install -r requirements.txt
streamlit run kap_risk_app.py
```

Tarayıcı `http://localhost:8501` adresinde açılır. İnternet bağlantısı gerekir; veriler canlı çekilir.

## Streamlit Community Cloud ile ücretsiz yayınlama

1. Bu depoyu GitHub'a itin (public).
2. [share.streamlit.io](https://share.streamlit.io) → **New app** → depo/dal seçin, ana dosya: `kap_risk_app.py`.
3. Birkaç dakika içinde `https://<kullanici>-<repo>.streamlit.app` biçiminde halka açık bir bağlantı alırsınız.

## Mimari (özet)

KAP'ın yeni Next.js arayüzü, sorgu sonuçlarını sunucu tarafında render edilen sayfanın içine gömülü
JSON (flight payload) olarak taşır. Uygulama bu yükü ayrıştırır; tarih filtresi sunucu tarafında
desteklenmediğinden sorgular yıl bazında yapılır (`&yr=`) ve seçilen aralık yerelde uygulanır.
Derin modda risk adayı bildirimlerin `Bildirim/{id}` detay sayfalarından tam metin çıkarılıp
sınıflandırma bu metin üzerinden yapılır (ör. derecelendirme notunun yönü).

## Sınırlamalar ve yasal not

- PDF/Excel ekleri taranmaz; kural tabanlı tespit nadir kalıpları kaçırabilir.
- KAP kısa aralıklı yoğun istekleri geçici olarak kısıtlayabilir; uygulama bunu algılar,
  bekler ve yeniden dener, veri alınamayan şirketleri açıkça işaretler.
- Bu araç **yatırım tavsiyesi değildir**; veriler KAP'ın kamuya açık bildirimlerinden derlenir,
  karar öncesi bildirim asıllarını doğrulayınız.
