# Değişiklik Günlüğü

## 3.0.0 — 2026-10-19 — İngilizce arayüz

- **İngilizce dil desteği:** Platform365 ana ekranı, Present ve EasyWrite'ın tüm arayüzü (menüler, iletişim kutuları, bildirimler, komut paleti, kısayol listeleri, masaüstü uygulama menüsü) Türkçe ve İngilizce
- **İlk açılışta dil seçimi:** Uygulama ilk açıldığında Türkçe / English sorulur; sonradan uygulama değiştiriciden (▦), ana ekrandaki 🌐 düğmesinden, Ctrl+K komut paletinden veya masaüstünde Görünüm → Dil menüsünden değiştirilebilir
- İngilizce seçildiğinde yeni slaytlar, şablonlar (mektup, özgeçmiş, toplantı notu, rapor, makale, proje planı, günlük), tarih/sayı biçimleri, dikte, canlı altyazı ve sesli okuma da İngilizce olur; EasyWrite okunabilirlik analizi Flesch puanı ve İngilizce dolgu kelimeleriyle çalışır
- **Platform365 hakkında:** sürüm, yayıncı, lisans, çalışma ortamı, bağlantılar ve sürüm notları — ana ekranda ayrı bölüm, her uygulamada ▦ menüsünden ve masaüstünde Yardım menüsünden pencere olarak
- **Yenilenen ana ekran:** sürüm rozeti, büyük karşılama alanı ve dosya sayıları, açıklamalı uygulama kartları, şablonlarla **hızlı başlangıç**, yeni özellik kartları ve düzenlenmiş veri / yedekleme bölümü
- Kullanıcı içeriği (slaytlar, belge metni, dosya adları) arayüz dili değişince asla çevrilmez

## 2.1.0 — 2026-10-17 — Masaüstü uygulaması

- **Windows (.exe) ve macOS (.dmg) masaüstü uygulaması** (Electron): kendi penceresi, Türkçe menü, çevrimdışı çalışma, .pptx / .present / .docx / .easywrite / .md dosya ilişkilendirmeleri (çift tıklayınca doğru uygulamada açılır), tek pencere
- **GitHub Releases:** `v*` etiketiyle Windows ve macOS paketleri otomatik derlenip yayımlanır
- **Web sürümü:** ilk açılışta masaüstü uygulamasını indirme penceresi; “⤓ İndir” düğmesi her zaman erişilebilir
- PWA (service worker ve web uygulaması bildirimi) kaldırıldı; önceki sürümden kalan önbellek otomatik temizlenir

## 2.0.0 — 2026-10-16 — Platform365

Present ve EasyWrite tek pakette birleşti: **Platform365**.

- **Platform365 ana ekranı:** iki uygulamanın dosyaları tek listede, içerikte arama, dosyayı doğru uygulamada açma, tam yedekleme / geri yükleme, depolama bilgisi
- **Ortak çekirdek:** uygulama değiştirici, Ctrl+K komut paleti, sistem/aydınlık/karanlık tema, otomatik ve adlandırılmış sürüm geçmişi, uygulamalar arası aktarım
- **EasyWrite (yeni):** Word uyumlu kelime işlemci — .docx/.easywrite/.md/.html/.txt okuma ve yazma, şablonlar, “/” komut menüsü, Markdown kısayolları, Türkçe okunabilirlik (Ateşman) ve yazım analizi, satır içi hesap, akıllı yazım, odak/daktilo modu, yazma hedefi, Pomodoro, dikte, sesli okuma, sürüm karşılaştırma, yapılacaklar listesi, bilgi kutuları, açılır bölümler, canlı içindekiler, bul/değiştir, Türkçe metin araçları, kağıt temaları
- **Present:** metinden sunum, EasyWrite belgesinden slayt / belgeye dönüştürme, sunum sağlık kontrolü, sunumda kalem/vurgulayıcı/lazer/büyüteç/ızgara/siyah ekran, canlı altyazı, prova modu, ilerleme çubuğu ve süre, tek dosya web sunumu (.html), ana hat (.md) dışa aktarma, emoji ekleme
- Present dosyaları `present/` klasörüne taşındı; eski bağlantılar ve kayıtlı sunumlar korunur

## 1.0.0 — 2026-10-04

İlk sürüm.

- Sunum oluşturma, açma, kopyalama ve silme (dosyalar cihazda IndexedDB'de saklanır)
- Slayt düzenleri: başlık, başlık ve içerik, bölüm başlığı, iki sütun, boş
- Metin kutusu, dikdörtgen, elips, hazır şekiller (ok, yıldız, kalp…), tablo ve görsel ekleme; sürükleme, boyutlandırma, döndürme ve hizalama kılavuzları
- Zengin metin: kelime düzeyinde kalın/italik/altı çizili, paragraf içi satır sonu (Shift+Enter)
- Yazı tipi, boyut, renk, kalın/italik/altı çizili, yatay ve dikey hizalama, dolgu ve kenarlık
- 6 hazır tema, slayt arka planı, konuşmacı notları
- Slayt ekleme, çoğaltma, silme ve sürükle-bırak ile sıralama
- Geri al / yinele, kopyala / yapıştır ve klavye kısayolları
- Tam ekran sunum modu (klavye, tıklama ve kaydırma ile gezinme)
- İndirme: Present dosyası (.present) — içeride geçerli bir PowerPoint paketi, PowerPoint/Keynote/Google Slaytlar/LibreOffice ile açılır; Present'e kayıpsız geri yüklenir. İsteğe bağlı .pptx uzantısıyla indirme; yazdırma / PDF
- .present ve PowerPoint (.pptx) dosyalarını açma (düğme veya sürükle-bırak); biçimli metin, hazır şekiller, serbest çizimler, tablolar, SmartArt, geçişler, doku dolguları, tema stilleri, kırpılmış görseller ve 4:3 dahil slayt boyutları korunur
- PowerPoint dosyası yazma: slayt içeriği doğrudan Office Open XML olarak üretilir (onarım uyarısı yok)
- PWA: yüklenebilir, çevrimdışı çalışır, yüklü uygulamada .present ve .pptx dosyalarını açar
