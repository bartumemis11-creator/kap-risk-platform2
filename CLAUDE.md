# Proje: KAP Risk İzleme ve Erken Uyarı Platformu

> Claude Code için proje anayasası. Her oturumda otomatik okunur. Türkçe yaz.
> Ayrıntılı Claude Code teknik rehberi: @docs/claude/GUNCEL_OZELLIKLER.md

## Ne yapar

KAP (Kamuyu Aydınlatma Platformu) bildirimlerini tarayıp BIST şirketlerinin kötüye gidişine
işaret eden sinyalleri tespit eder, ağırlıklı 0–100 risk skoru + A–E notu üretir, biçimlendirilmiş
Excel/CSV raporu sunar. **Yatırım tavsiyesi değildir.**

## Tech Stack

- Python 3.13 (bkz. `.python-version`)
- Streamlit (UI) + pandas + openpyxl (Excel) + altair (grafik) + requests (KAP/haber çekimi)
- Veri KAP'ın Next.js arayüzündeki gömülü JSON (flight payload) ayrıştırılarak canlı çekilir

## Klasör / dosya yapısı

```
kap_risk_app.py        # Streamlit UI (giriş noktası)
kap_risk_tarama.py     # Çekirdek: tarama, sınıflandırma, skorlama + argparse CLI
test_core.py           # Çekirdek birim testleri (pytest)
test_golden.py         # 24 vakalık altın regresyon seti
requirements.txt       # Çalışma bağımlılıkları
docs/claude/           # Claude Code teknik rehberleri
```

## Komutlar

| Amaç | Komut |
|---|---|
| Dev server (UI) | `streamlit run kap_risk_app.py` |
| Testler | `python -m pytest -q` |
| Tek test | `python -m pytest test_core.py::test_rating_direction -q` |
| CLI tarama | `python kap_risk_tarama.py --help` |
| Bağımlılık kur | `pip install -r requirements.txt` |

> `pytest` requirements.txt'te yok; testler için `pip install pytest` gerekir.

## Kurallar (DO)

- **Türkçe** commit mesajı ve kullanıcıya dönük metin yaz.
- Yeni risk kalıbı/kategori eklerken **hem `test_core.py` hem `test_golden.py`** güncellensin.
- Veri çekiminde KAP dostu ol: hız sınırlayıcı + devre kesici + yeniden deneme mantığını koru.
- Sessiz veri kaybından kaçın — veri alınamayan şirket açıkça işaretlenmeli.
- Model ID gerektiğinde güncel olanı kullan: `claude-opus-4-8` (varsayılan), `claude-sonnet-5`, `claude-haiku-4-5`.

## Yapma (DON'T)

- ❌ Testleri kırıp PR açma; çekirdek mantığı testsiz değiştirme.
- ❌ Skor ağırlıklarını (kategori → ağırlık) gerekçesiz değiştirme — regresyon setini bozar.
- ❌ Kişisel/yerel çıktıları (`kap_risk_izleme_gecmis.json`, `KAP_Risk_Raporu_*.xlsx`, `*.csv`) commit etme.
- ❌ Eski model ID'si (`claude-opus-4-7`, `claude-sonnet-4-6`) yazma.

## Tamamlanma Kriterleri (DoD)

- ✓ `python -m pytest -q` tümü yeşil (çekirdek + altın set)
- ✓ Değişiklik UI'da gözlemlenebiliyorsa `streamlit run` ile doğrulandı
- ✓ Anlamlı, Türkçe, açıklayıcı commit mesajı

## Hata Protokolü

3 denemeden sonra başarısız olursan: DUR, hata raporunu yaz, "İNSAN_GEREKLİ: [açıklama]" de ve bekle.
