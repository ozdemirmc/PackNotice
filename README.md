# Pack Mailler WEB

Bakım paketleri için otomatik mail hazırlama aracı. Bu proje, bir Outlook Web Add-in (Eklenti) olarak çalışır.

## Kurulum ve Dağıtım

Bu eklenti GitHub Pages üzerinde barındırılmaktadır.

1. **Manifest Yükleme:** `manifest.xml` dosyasını Outlook'un "Özel Eklenti Ekle" kısmından yükleyin.
2. **Kullanım:** Outlook'ta yeni bir mail oluştururken veya bir maili yanıtlarken "Pack Mailler" butonuna tıklayarak task pane'i açabilirsiniz.

## Dosya Yapısı

- `index.html`: Eklentinin ana arayüzü.
- `app.js`: Mail hazırlama ve Office.js entegrasyon mantığı.
- `style.css`: Görsel tasarım dosyası.
- `settings.js`: Kullanıcı ayarları ve yerel depolama yönetimi.
- `manifest.xml`: Outlook için yapılandırma dosyası.
- `assets/`: Eklenti ikonları.

---
*Hazırlayan: ozdemirmc*
