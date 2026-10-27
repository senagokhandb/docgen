# DOCGEN skill paketleri — OPPY entegrasyon notu

Sürüm: docgen 1.1.0 · 27 Ekim 2026 · Hazırlayan: Sena Gökhan

## Paketin içeriği

| Dosya | Ne |
|---|---|
| `docx.skill` | Word: oluşturma, yüklenen belgeyi yerinde düzenleme, kullanıcı şablonu (.dotx) |
| `pptx.skill` | PowerPoint: temalı deste, yerinde düzenleme, kullanıcı şablonu (.potx) |
| `xlsx.skill` | Excel: formüllü çalışma kitabı, yerinde düzenleme, satır/sütun işlemleri, yeniden hesap |
| `pdf.skill` | PDF: HTML'den üretim, metin ve sayfa düzenleme, birleştirme/bölme/filigran |
| `manifest.json` | Skill listesi ve sürüm |
| `QWEN.md` | Sistem kuralları (ajanın sistem prompt'u) |

Her `.skill` dosyası bir zip arşivi: içinde `<ad>/SKILL.md` (iş akışı) ve
`<ad>/REFERENCE.md` (teknik ayrıntı) var.

## Önemli: skill'ler tek başına çalışmaz

Skill'ler ajana bir işi *nasıl* yapacağını anlatan talimatlardır. Asıl işi
docgen çalışma ortamındaki betikler ve servisler yapar. OPPY'de çalışmaları
için şunlar gerekir:

1. **Kod çalıştırabilen bir ajan.** Ajan kendi Python/Node betiklerini
   yazıp shell'de çalıştırır. Yalnızca sohbet eden bir model bu skill'leri
   kullanamaz.
2. **`QWEN.md` sistem prompt'u olarak.** Skill'ler bu kurallara atıf yapar
   (ör. "QWEN.md › Themes"). Düzenleme kuralları skill'lerin önüne geçer.
   Dosyadaki `/home/agent/.qwen/skills/` yolunu kendi ortamınızdaki skill
   klasörüyle değiştirin.
3. **docgen çalışma ortamı** (skill'ler bu yolları doğrudan çağırır):
   - `/opt/docgen/scripts/`: doğrulama ve düzenleme betikleri
     (`validate.py`, `diff_check.py`, `readback.py`, `render_check.py`,
     `find_target.py`, `pdf_text.py`, `pdf_insert.py`, `pdf_pages.py`,
     `xlsx_recalc.py`, `xlsx_struct.py`, `repair_pptx.py`,
     `build_themed_deck.js`, …)
   - `/opt/docgen/lib/docgen`: Python kütüphanesi (`PYTHONPATH=/opt/docgen/lib`)
   - `/opt/docgen/themes/`: tema token'ları, CSS ön ayarları, şablonlar
   - `/opt/docgen/schemas/`: ECMA-376 XSD şemaları (`validate.py` için)
4. **Yazılımlar:** Python 3.12 (`requirements.txt`), Node 22 + `docx` ve
   `pptxgenjs` (`NODE_PATH`), LibreOffice headless (Calc/Writer/Impress) ve
   Carlito, Caladea, Liberation, DejaVu fontları.
5. **Gotenberg 8**, adresi `GOTENBERG_URL` ortam değişkeninde.
6. **Dosya düzeni:** `/workspace/uploads/` (kullanıcı dosyaları, salt
   okunur), `/workspace/scratch/` (çalışma alanı), `/workspace/outputs/`
   (teslim edilen dosyalar).

**En kolay yol:** docgen deposundaki `docgen-agent` Docker imajını ve
Gotenberg servisini (`docker-compose.yml`) olduğu gibi kullanmak. Yukarıdaki
ortamın tamamı bu imajda hazır. Depoya erişim için Sena'ya yazın.

## Sohbet arayüzüne bağlarken

Takip isteklerinde ("başlığı kırmızı yap") önceki mesajın teslim ettiği
dosyayı `uploads/` altına koyun. Önceki mesajların `scratch/` klasörlerini
de `/workspace/previous/<n>/` altına salt okunur bağlayın: HTML'den
üretilmiş bir PDF'te renk, tablo ya da bölüm değişikliği ancak kaynak HTML
düzenlenerek yapılabilir. Örnek uygulama depoda: `ui/server.py`.

## Bilinen sorun

PDF'ler şu an Letter boyutunda çıkıyor. HTML'de A4 tanımlı, ama Gotenberg
çağrısına `preferCssPageSize=true` eklenmediği için bu ayar yok sayılıyor.
