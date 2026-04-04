# PackNotice

Basılıp hazır edilen bakım paketleri için kurumsal formatta mail hazırlama aracı.

## Amaç

Bu proje, uçak bakım paketlerinin tamamlanmasının ardından ilgili birimlere (Bay 1-2-3) gönderilen bilgilendirme maillerini standart bir formata sokmak için geliştirilmiştir.

Ekip içindeki farklı üslupları ve formatları ortadan kaldırarak kurumsal bir iletişim dili oluşturmayı hedefler.

## Özellikler

- **Hızlı Giriş**: A/C, Bakım Adı ve Tarih bilgilerini kolayca girin.
- **Dinamik Bay Seçimi**: Mailin gönderileceği alıcıyı tek tıkla belirleyin.
- **Bakım Tipi Seçimi**: Planlı, Plansız veya Periyodik bakıma göre otomatik zimmet tabloları oluşturur.
- **Zimmet Modu**: BİRİM veya PLANNER modunda çalışır — her mod için farklı mail içeriği üretilir.
- **Tarih Kısıtlaması**: Geçmiş tarih seçimi engellenir. Seçilen tarih bugünse mail içeriğinde **(BUGÜN)** olarak kırmızı vurgulanır.
- **Gönderici Kontrolü**: Mailin TT-UBB(SAW)-BAKIMHAZIRLIK hesabından gönderilip gönderilmediğini denetler.
- **Ayarlar Yönetimi**: BAY-1/2/3 alıcıları ve CC listesi kalıcı olarak kaydedilir. Ayarlar varsayılana döndürülebilir.
- **Outlook Entegrasyonu**: Doğrudan Outlook Taskpane üzerinden çalışır.

## Kurulum ve Kullanım

1. Proje dosyalarını GitHub Pages üzerinden host edin.
2. `manifest.xml` dosyasındaki URL kısımlarını kendi host adresinizle güncelleyin.
3. XML dosyasını Outlook'ta eklenti olarak tanıtın.

## Versiyon

v1.5 — BAKIM PLANLAMA ŞEFLİĞİ (SAW) / BAKIM HAZIRLIK BİRİMİ
