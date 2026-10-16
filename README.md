# Platform365

**Sürüm 2.0.0** — Tarayıcıda çalışan, yüklenebilir (PWA) ve çevrimdışı kullanılabilen ofis paketi. İki uygulamayı tek çatıda birleştirir:

| Uygulama | Ne yapar | Uyumluluk |
| --- | --- | --- |
| **Present** | Sunum hazırlama ve sunma | PowerPoint (.pptx), kendi biçimi .present |
| **EasyWrite** | Belge / kelime işlemci | Word (.docx), kendi biçimi .easywrite, Markdown, HTML, TXT |

Kök adres **Platform365 ana ekranıdır**: iki uygulamanın son dosyaları tek listede, içerikte arama, dosya bırakınca doğru uygulamada açma, tam yedekleme. Sağ üstteki ▦ düğmesiyle (uygulama değiştirici) uygulamalar arasında geçilir; **Ctrl+K** her yerde komut paletini açar. Hesap, sunucu ya da bulut yoktur — her şey cihazda (IndexedDB) saklanır.

## PowerPoint ve Word'de olmayan özellikler

### Platform365 (her iki uygulamada)

- **Belge ⇄ sunum dönüşümü:** EasyWrite belgesi tek tıkla Present sunumu olur (Başlık 1/2 → slayt, paragraf ve maddeler → slayt maddesi); sunum da slaytlar, maddeler ve konuşmacı notlarıyla düzenlenebilir belgeye dönüşür.
- **Komut paleti (Ctrl+K):** Tüm komutlara yazarak ulaşma, Türkçe karakter duyarsız arama.
- **Sürüm geçmişi:** Dosya her açıldığında (30 dakikada bir) ve riskli işlemlerden önce otomatik sürüm; adlandırılmış sürümler; önizleme ve geri yükleme (geri yükleme de Ctrl+Z ile geri alınabilir).
- **Ortak arama:** Ana ekrandan tüm sunum ve belgelerin başlık **ve içeriğinde** arama.
- **Tek tıkla tam yedek:** Tüm sunumlar, belgeler ve sürümler tek `.p365` dosyasında; başka cihaza geri yükleme.
- **Akıllı dosya açma:** Ana ekrana bırakılan .pptx/.docx/.md… dosyası doğru uygulamada açılır; EasyWrite'a .pptx bırakılırsa Present'e yönlendirilir.
- **Tema:** Sistem / aydınlık / karanlık; tüm uygulamalarda eşzamanlı.

### Present

- **Metinden sunum:** Ana hat (Markdown veya girintili metin) yapıştırın; başlıklar slayt, maddeler (iç içe seviyelerle) madde, `>` satırları konuşmacı notu olur. 7 maddeden uzun slaytlar otomatik bölünür, yazı boyutu içeriğe göre ayarlanır.
- **Sunum sağlık kontrolü:** WCAG kontrast oranı, arkadan okunamayacak küçük yazı, kalabalık slayt (kelime sayısı), slayttan taşan/dışarıda kalan öğe, boş metin kutusu, yazı tipi karmaşası — 100 üzerinden puan ve tek tıkla sorunlu öğeye gitme.
- **Sunarken araçlar:** Kalem (D, tekrar basınca renk değişir), vurgulayıcı (H), lazer işaretçi (L), büyüteç (Z — tıklanan bölgeyi büyütür), çizimleri silme (E), siyah/beyaz ekran (B/W), slayt ızgarası (G), sayı + Enter ile slayta atlama, ilerleme çubuğu ve geçen süre.
- **Canlı altyazı (C):** Sunum sırasında konuşma, Türkçe altyazı olarak ekranın altında görünür (Chrome/Edge).
- **Prova modu:** Slayt başına süre ölçümü, uzun kalan slaytların işaretlenmesi, sürelerin konuşmacı notlarına yazılması.
- **Tek dosya web sunumu (.html):** Programsız açılan, klavye/dokunmatik ile gezilen, tek dosyalık sunum.
- **Ana hat dışa aktarma (.md)** ve **emoji/ikon ekleme**.

### EasyWrite

- **“/” komut menüsü:** Başlık, liste, yapılacaklar, alıntı, kod, bilgi/uyarı/başarı/tehlike kutusu, açılır bölüm, tablo, görsel, sayfa sonu, içindekiler, tarih, saat, emoji…
- **Markdown ile yazma:** `#` başlık, `-` madde, `1.` numaralı, `[]` görev, `>` alıntı, `**kalın**`, `*italik*`, `~~çizili~~`, `==vurgu==`, `` `kod` ``, `---` çizgi, ```` ``` ```` kod bloğu, `[toc]` içindekiler. Markdown metni yapıştırınca biçimli gelir.
- **Türkçe okunabilirlik analizi:** Ateşman puanı, okuma ve konuşma süresi, uzun cümle (25+ kelime) ve dolgu kelime (aslında, yani, sanki…) vurgulama, sık tekrarlanan kelimeleri tıklayıp belgede işaretleme — belge değiştirilmeden (CSS Highlight API).
- **Satır içi hesap:** `1250*12%=` yazınca sonucu (`=150`) ekler; Türkçe ondalık virgülünü anlar.
- **Akıllı yazım:** “Türkçe tırnak”, — uzun tire, …, ©, ®, ™, →, ←, ⇒, ½.
- **Odak modu** (yalnızca yazılan paragraf parlak) ve **daktilo kaydırma** (imleç hep ortada).
- **Yazma hedefi** (kelime hedefi, ilerleme, oturum sayacı) ve **Pomodoro** zamanlayıcı.
- **Dikte** (konuşarak yazma) ve **sesli okuma** (Türkçe).
- **Sürüm karşılaştırma:** Eski sürümle şimdiki hâl arasında eklenen/silinen kelimeler.
- **Etkileşimli yapılacaklar listesi**, renkli **bilgi kutuları** (simgeye tıklayınca tür değişir), **açılır bölümler**, tıklanabilir **canlı içindekiler** ve **ana hat paneli**.
- **Türkçe metin araçları:** BÜYÜK/küçük harf (i/İ, ı/I kurallarıyla), Başlık Durumu, Cümle durumu, metin temizleme, liste/paragraf sıralama (Türkçe alfabe).
- **Paragraf taşıma** (Alt+↑/↓), tabloda Tab ile hücre gezme (sonda yeni satır), tablo/görsel bağlam araç çubuğu.
- **Kağıt temaları:** Beyaz, sepya, nane, gece (yazdırırken beyaz).
- **Şablonlar:** Mektup, özgeçmiş, toplantı notu, rapor, makale, proje planı, günlük.

## Dosya biçimleri

- **.present** — İçeride geçerli bir PowerPoint (.pptx) paketi + kayıpsız Present verisi (`present/deck.json`).
- **.easywrite** — İçeride geçerli bir Word (.docx) paketi + kayıpsız EasyWrite verisi (`easywrite/doc.json`). Uzantıyı `.docx` yapmak Word'de açmak için yeterlidir.
- **.docx** — EasyWrite başlıkları, kalın/italik/altı/üstü çizili, renk, vurgu, yazı boyutu ve tipi, üst/alt simge, bağlantılar, madde/numaralı (iç içe) listeler, yapılacaklar (☐/☒), tablolar (başlık satırı, birleştirilmiş sütunlar), görseller, alıntı, kod, bilgi kutuları, ayraç, sayfa sonu, sayfa boyutu ve kenar boşluklarıyla yazar; Word dosyalarını aynı öğelerle okur.
- **.md / .html / .txt** — Dışa ve içe aktarma.
- **.p365** — Platform365 tam yedeği (JSON).

### PowerPoint uyumluluğu (Present)

#### PowerPoint → Present (açma)

- **Metin:** Kelime düzeyinde biçim (kalın, italik, altı çizili, üstü çizili, renk, boyut, yazı tipi, büyük harf), paragraf hizalaması, madde işaretleri (Wingdings dahil) ve numaralandırma, girinti seviyeleri, satır ve paragraf aralıkları, iç boşluklar, dikey hizalama, PowerPoint'in kaydettiği otomatik sığdırma oranı.
- **Şekiller:** 100'e yakın hazır şekil (oklar, yıldızlar, akış şeması, pasta/yay, bulut, parantez…), serbest çizimler (custGeom), bağlayıcı çizgiler ve ok uçları, kesikli çizgiler, döndürme ve yansıtma.
- **Dolgular:** Düz renk, saydamlık, doğrusal ve radyal geçişler, resim/doku dolgusu (döşemeli dahil), tema stilleri (fillRef/lnRef), tema renkleri ve renk değiştiricileri (tint, shade, lumMod, satMod…).
- **Görseller:** Kırpma, yuvarlak/elips çerçeve, kenarlık.
- **Tablolar:** Gerçek tablo olarak; birleştirilmiş hücreler, hücre dolguları, tablo stilleri (başlık satırı, bantlı satırlar…), kenarlıklar.
- **SmartArt:** PowerPoint'in kaydettiği çizimden şekil ve metin olarak.
- **Diğer:** Gruplar, ana slayt ve düzen öğeleri, yer tutucu kalıtımı, arka planlar (renk, geçiş, resim), slayt boyutu (16:9, 4:3…), konuşmacı notları.
- **Aktarılmayanlar:** Grafikler (yerine bilgi kutusu konur), video/ses, animasyon ve geçişler, gölge/parlama efektleri, EMF/WMF görseller.

#### Present → PowerPoint (indirme)

Slayt içeriği PowerPoint'in beklediği yapıda (Office Open XML) doğrudan yazılır: metin kutuları kelime düzeyinde biçimiyle, hazır şekiller ve serbest çizimler kendi geometrisiyle, tablolar gerçek PowerPoint tablosu olarak, geçişler ve doku dolguları PowerPoint dolgusu olarak kaydedilir. Böylece dosya PowerPoint'te "onarım" uyarısı vermeden açılır ve düzenlenebilir kalır.

## Klavye kısayolları

| Kısayol | İşlem |
| --- | --- |
| Ctrl+K | Komut paleti (her yerde) |
| Ctrl+Z / Ctrl+Y | Geri al / Yinele |
| Ctrl+S | Kaydet |
| **Present** | |
| Çift tıklama / Enter | Metni veya tablo hücresini düzenle |
| Shift+Enter (düzenlerken) | Paragraf içinde satır sonu |
| Ctrl+B / Ctrl+I / Ctrl+U | Kalın / İtalik / Altı çizili |
| Ctrl+C / Ctrl+X / Ctrl+V, Ctrl+D | Kopyala / Kes / Yapıştır, Çoğalt |
| Ok tuşları (Shift ile 10px) | Öğeyi kaydır |
| Alt + sürükleme / Shift + köşe | Kılavuzsuz taşı / Oranı koru |
| F5 / Shift+F5 | Baştan sun / Geçerli slayttan sun |
| Sunumda D, H, L, Z, E | Kalem, vurgulayıcı, lazer, büyüteç, çizimleri sil |
| Sunumda B, W, G, C, T | Siyah ekran, beyaz ekran, ızgara, canlı altyazı, süreyi sıfırla |
| **EasyWrite** | |
| / | Komut menüsü |
| Ctrl+F / Ctrl+H | Bul / Değiştir |
| Ctrl+Shift+F | Odak modu |
| Ctrl+Alt+1…3 / Ctrl+Alt+0 | Başlık 1–3 / Normal |
| Ctrl+Shift+7 / 8 / 9 | Numaralı / madde / yapılacaklar listesi |
| Ctrl+Enter | Sayfa sonu |
| Alt+↑ / Alt+↓ | Paragrafı taşı |
| Ctrl+Shift+V | Biçimsiz yapıştır |
| Ctrl+P | Yazdır / PDF |

## Dosya yapısı

```
index.html, hub.js, hub.css   Platform365 ana ekranı
shared/platform.js            Ortak çekirdek: tema, uygulama değiştirici, komut paleti, iletişim kutuları,
                              uygulamalar arası gelen kutusu, sürüm geçmişi, konuşma tanıma
shared/platform.css           Ortak stiller ve renk değişkenleri (aydınlık / karanlık)
shared/convert.js             Markdown ⇄ HTML, güvenli HTML temizleme, belge/metin → sunum ana hattı
present/                      Present: index.html, app.js (düzenleyici), features.js (Platform365 eklentileri),
                              shapes.js, pptx-import.js, pptx-export.js, styles.css
easywrite/                    EasyWrite: index.html, app.js (düzenleyici), docx.js (Word okuma/yazma), styles.css
sw.js                         Service worker (tüm platform için çevrimdışı önbellek)
manifest.webmanifest          PWA bildirimi (dosya ilişkilendirmeleri: .present, .pptx, .easywrite, .docx, .md)
icons/                        Uygulama ikonları
vendor/                       PptxGenJS 3.12.0 (MIT, JSZip içerir)
```

Çalıştırmak için herhangi bir statik sunucu yeterlidir (ör. `npm start` → http://localhost:8080). Eski Present adresleri (`/#/d/…`) otomatik olarak `present/` altına yönlendirilir; mevcut sunumlar aynı tarayıcı deposunda kaldığı için kaybolmaz.

## Sürüm güncelleme

Yeni sürümde `present/app.js` ve `easywrite/app.js` içindeki `APP_VERSION`, `shared/platform.js` içindeki `VERSION`, `sw.js` içindeki `VERSION` (ve dosyalar her değiştiğinde `BUILD`), `manifest.webmanifest` ve `package.json` sürümlerini birlikte güncelleyin.

## Lisans

MIT. [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (MIT) ve içerdiği [JSZip](https://github.com/Stuk/jszip) (MIT) kullanılır; lisansları `vendor/` klasöründedir.
