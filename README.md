# DOCGEN

Docker tabanlı, LLM ajan destekli belge üretim sistemi. Doğal dilde verilen
bir istek; DOCX, PPTX, XLSX veya PDF olarak üretilmiş (veya düzenlenmiş),
**doğrulanmış** ve kullanıma hazır bir dosya olarak teslim edilir.

```
kullanıcı istegi ──> docker compose ──> agent (Qwen Code, non-root sandbox)
                                                    │   │
                                                    │   └──> model endpoint (Gemma, .env)
                                                    └──> Gotenberg (HTML/ofis→PDF, PDF işlemleri)
```

---

## 1. Mimari

### 1.1 Servisler (docker-compose.yml)

| Servis | Görüntü | Görevi |
|---|---|---|
| **agent** | `docgen-agent:latest` (Dockerfile ile derlenir) | Belge üretim ajanı: Qwen Code CLI, Python/Node motorları, skills, doğrulama scriptleri, temalar, şablonlar, XSD şemaları. Non-root `agent` (uid 1000) kullanıcı olarak çalışır. |
| **gotenberg** | `gotenberg/gotenberg:8` | Render ve PDF motoru: Chromium (HTML→PDF), LibreOffice (ofis→PDF), pdfengines (merge/split/watermark/stamp/rotate). Host portu açılmaz; yalnızca iç ağdan erişilir. |

### 1.2 Agent konteynerinin içindekileri

| Katman | İçerik |
|---|---|
| **Ajan çekirdeği** | Qwen Code CLI v0.26.1 — istemi headless modda (`--yolo -o json`) çalıştırır; kod yazma, shell komutu, dosya okuma/yazma yetkileri tam. Headless modun gömülü "soru sorma, varsayıp bitir" cümlesi imaj derlenirken `patch_qwen_prompt.py` ile değiştirilir: her çalışma bir sohbet turudur; ajan yalnızca kullanıcının kararı olmadan güvenle ilerleyemeyeceği yerde (araç reddi, belirsiz hedef, onaysız yıkıcı değişiklik) sorar, gerisinde varsayımını söyleyip işi bitirir. |
| **Model** | `.env` dosyasındaki OpenAI-uyumlu endpoint (Gemma). Konteyner dışı, kendi inference servisiniz |
| **docgen kütüphanesi** | `/opt/docgen/lib/docgen` (`PYTHONPATH`) — yüklenen dosyayı **yerinde** düzenlemek için test edilmiş yapı taşları: `package`, `docx`, `pptx`, `xlsx`, `xlsx_struct`, `pdf`, `theme`, `compare`, `office`. Agent kendi script'inden çağırır. |
| **Python motoru** | office_oxide (hızlı okuma, yeni dosya), python-docx / python-pptx (kullanıcı şablonundan üretim), openpyxl (yeni çalışma kitabı), pdfplumber, pypdfium2 (PDF sayfa + metin düzenleme), lxml, requests, defusedxml, Pillow |
| **LibreOffice (headless)** | yalnızca hesap motoru ve dönüştürücü: formül sonuçlarının hesaplanması (`xlsx_recalc.py`), `.doc/.xls/.ppt` dönüşümü (`convert_legacy.py`) |
| **Fontlar** | Gotenberg ile aynı aileler (Caladea, Carlito, Liberation, DejaVu) — PDF metni aynı fontla yeniden çizilebilsin diye |
| **Node motoru** | docx-js (docx), pptxgenjs — yeni dosya üretimi; `NODE_PATH=/opt/docgen/node_modules` |
| **Bilgi katmanı** | `/opt/docgen/skills/` (4 format skill + referanslar), `/opt/docgen/QWEN.md` (sistem kuralları), `/opt/docgen/themes/` (temalar + CSS + şablonlar), `/opt/docgen/schemas/` (ECMA-376 XSD seti) |
| **Doğrulama araç seti** | `/opt/docgen/scripts/` — validate (`--baseline`), diff_check, find_target, readback (`--format`), render_check, xlsx_recalc, xlsx_struct, pdf_text, pdf_pages, convert_legacy, recalc_check, visual_qa, repair_pptx, build_templates, build_themed_deck, pack_skill |

### 1.3 Dosya modeli (sandbox kuralları)

| Yol | Erişim |
|---|---|
| `/workspace/uploads/` (host: `data/uploads/`) | salt okunur — kullanıcı kaynakları (compose'da `:ro` mount; orijinal her zaman korunur, "geri al" = orijinal) |
| `/workspace/scratch/` (host: `data/scratch/`) | okuma/yazma — ara dosyalar, scriptler |
| `/workspace/outputs/` (host: `data/outputs/`) | **teslimatlar yalnızca burada** |
| `/opt/docgen/` | salt okunur — sistem |

Konteyner her istekte tek seferlik çalışır (`docker compose run --rm agent`),
çıkışta `/workspace` bind-mount üzerinden hosta geri döner; state birikmez.

---

## 2. Bir istek nasıl işlenir (akış)

```
1. docker compose run --rm agent "<istek>"
2. entrypoint.sh
   - skill + sistem kurallarını ~/.qwen/ altına kopyalar
   - Gotenberg health poll (en fazla 60 sn)
   - qwen "$REQUEST" --model $MODEL_NAME --openai-base-url $MODEL_ENDPOINT --yolo -o json
3. Ajan (QWEN.md kuralları)
   a. İlgili format skill'ini ÖNCE OKUR (zorunlu)
   b. YENİ belge: tema (ya da kullanıcı şablonu) seçer, motoru seçer, kodu yazar, üretir
      YÜKLENEN belge: find_target.py ile hedefi bulur (birden çok aday -> dosya yazmaz,
      adayları raporlar), docgen kütüphanesiyle YERİNDE düzenler (baştan üretmez)
   c. DOĞRULAMA DÖNGÜSÜ:
        validate.py [--baseline orijinal] -> yapısal + ECMA-376 XSD (yalnızca YENİ hatalar)
        diff_check.py orijinal sonuç --allow-...  -> istenen dışında bir şey değiştiyse FAIL
        xlsx_recalc.py -> formül sonuçlarını LibreOffice ile hesaplayıp dosyaya yazar
        readback.py [--format] -> içerik / biçim geri okuma
        render_check.py -> LibreOffice (Gotenberg) dosyayı sayfalara dökebiliyor mu
      Başarısız adım: düzelt + tekrar (maks. 3 tur, sonra dürüst rapor)
   d. Son dosyayı /workspace/outputs/ koyar; neyin değiştiğini ve neyin
      yapılamadığını (nedeniyle) raporlar; outputs/manifest.json'a kayıt düşer
```

Görsel QA **varsayılan olarak kapalıdır**; istekte `--visual-qa` varsa
`visual_qa.py` devreye girer (Bölüm 5.11).

---

## 3. Motor matrisi

| İhtiyaç | Araç | Dil |
|---|---|---|
| Ofis dosyası okuma (markdown/HTML/düz metin) | office_oxide `Document` | Python |
| **Yüklenen DOCX/PPTX/XLSX'i yerinde düzenleme** (metin, biçim, konum, tablo, bölüm, slayt, stil, tema) | docgen kütüphanesi (`docgen.docx/pptx/xlsx/theme`) | Python |
| XLSX satır/sütun ekle-sil, sıralama (referanslar her yerde kayar) | `docgen.xlsx_struct` / `xlsx_struct.py` | Python |
| XLSX formül sonuçlarını hesaplayıp dosyaya yazma | `xlsx_recalc.py` (LibreOffice hesap motoru) | Python |
| PDF metin (cümle, tablo hücresi, üst/alt bilgi), yeni paragraf ekleme ve sayfa düzenleme | `docgen.pdf` / `pdf_text.py`, `pdf_insert.py`, `pdf_pages.py` (pypdfium2) | Python |
| Kullanıcı şablonundan yeni belge/deste | `start_from_template` + python-docx / python-pptx | Python |
| `.doc/.xls/.ppt` → OOXML | `convert_legacy.py` (LibreOffice) | Python |
| Markdown'dan DOCX/PPTX/XLSX (hızlı yol, yeni dosya) | office_oxide `create_from_markdown` | Python |
| XLSX üretimi: hücre, formül, birleştirme, stil (yalnız kalın + arka plan) | office_oxide `XlsxWriter` | Python |
| Yeni XLSX: tam stil, grafik, doğrulama | openpyxl (yalnız yeni dosya; yüklenen dosyayı kaydederken formül sonuçlarını siler) | Python |
| DOCX tam kontrol (yeni dosya) | docx (docx-js) | Node |
| PPTX tam kontrol (yeni dosya) | pptxgenjs | Node |
| PDF metin/tablo çıkarma (QA geri-okuma) | pdfplumber | Python |
| XML parse + XSD doğrulama | lxml | Python |
| PDF→PNG render (görsel QA) | pypdfium2 + Pillow | Python |
| HTTP istemcisi (Gotenberg) | requests | Python |
| HTML→PDF | Gotenberg Chromium | servis |
| Ofis→PDF | Gotenberg LibreOffice | servis |
| PDF merge / watermark / stamp / rotate | Gotenberg pdfengines (`crop` diye bir rota yoktur) | servis |

`office_oxide.EditableDocument.replace_text` hedefli düzenleme için
kullanılmaz: her geçişi (kelime içindekiler dahil: "risk" → "riski")
değiştirir ve run sınırını aşan metinde biçimi yayar.

Tüm bağımlılıklar açık kaynak (MIT / BSD / Apache-2.0 / HPND / PSF);
lisanslar `requirements.txt` ve `package.json` içinde satır satır belirtilir.

### 3.1 Eleman seviyesinde stil kontrolü

Agent hazır script'lerle sınırlı değildir: her istek için kendi kodunu
`/workspace/scratch/` altına serbestçe yazar. Bu yüzden "şu yazıyı şu renk
yap" gibi bir istek doğrudan koda çevrilir ve yalnızca istenen
paragraf / run / hücre / slayt / şekil üzerinde uygulanır.

| Format | Yeni dosya | Yüklenen dosya (yerinde) | Kontrol edilenler |
|---|---|---|---|
| DOCX | docx-js `TextRun`, `Paragraph`, `TableCell` | `docgen.docx`: `set_run_props`, `set_paragraph_props`, `set_style_props`, `set_cell_shading` | renk, font, punto, kalın/italik/altı/üstü çizili, vurgu, hizalama, satır aralığı, girinti, hücre dolgusu; tek ifade / paragraf / stil (tüm başlıklar) kapsamında |
| PPTX | pptxgenjs `addText` (`x/y/w/h`, run dizileri) | `docgen.pptx`: `set_run_props`, `set_paragraph_props`, `set_geometry`, `scale`, `set_fill`, `set_background` | renk, font, punto, hizalama, **konum ve boyut**, dolgu, tek slayt arka planı, tablo hücresi |
| XLSX | openpyxl `Font`, `PatternFill`, `Alignment` | `docgen.xlsx`: `set_style`, `add_conditional_format`, `set_column_width` | font rengi/adı/punto, kalın/italik/altı çizili, dolgu, sayı biçimi, hizalama, kalıcı koşullu biçim |
| PDF | HTML + eleman bazlı CSS → Gotenberg | yalnız metin içeriği (`pdf_text.py`) | mevcut PDF metninin renk/font değişimi desteklenmez (kaynak dosya düzenlenir: önce değiştirmeden render edilip `diff_check` ile PDF'i birebir verdiği kanıtlanır, sonra değişiklik yapılıp yeniden render edilir) |

- **office_oxide stil motoru değildir:** `XlsxWriter` yalnız kalın + arka
  plan rengi verebilir; `PptxWriter` ve `EditableDocument` stil ayarlayamaz.
- Yardımcılar özellikleri şema sırasıyla yazar, run'ı tam karakter
  sınırından böler ve stil kalıtımını hesaba katar (stilden gelen kalınlığı
  kaldırmak için açıkça "kapalı" yazar).

### 3.2 Yüklenen dosyayı yerinde düzenleme (yetenek ve korumalar)

Yüklenen belge baştan üretilmez; yalnızca istenen eleman değişir. Aşağıdaki
yetenekler kabul testi listesindeki gerçek fixture'larla (Word/LibreOffice
ile üretilmiş belgeler) denenmiş ve `diff_check` + `validate --baseline`
ile doğrulanmıştır.

| Format | Yerinde yapılabilenler | Değişen paket parçası (örnek) |
|---|---|---|
| DOCX | tek ifade/geçiş değiştirme (bölüm/başlık bağlamıyla), run/paragraf/stil biçimi, tablo hücresi, satır ve **sütun** ekleme/silme, paragraf ekleme, bölüm silme, tema renkleri, şablondan yeni belge | yalnız `word/document.xml` (stil kapsamında `styles.xml`, tema için `theme1.xml`) |
| PPTX | madde/metin, run biçimi, **konum/boyut**, resim ölçekleme, dolgu, tek slayt arka planı, konuşmacı notu, tablo hücresi/biçimi/sütun ekleme, slayt silme/taşıma/çoğaltma/gizleme, tema rengi, şablondan yeni deste | yalnız ilgili `slideN.xml` / `notesSlideN.xml` (yapısal işlemlerde `presentation.xml` + ilişkiler) |
| XLSX | değer/formül (stil korunur), hücre biçimi, sayı biçimi, koşullu biçim, sütun genişliği, **satır/sütun ekleme-silme ve sıralama** (formüller, adlar, CF/DV, tablo, grafik serileri, çizim çapaları birlikte kayar), formül sonuçlarının yeniden hesaplanıp yazılması, `.xltx` şablon doldurma | yalnız ilgili sayfa(lar), `styles.xml`, tablo tanımı |
| PDF | cümle/paragraf metni (aynı font ailesi, sütun içinde satır kaydırma), tablo hücresi, üst/alt bilgi metni, sayfa silme/taşıma/ekleme/döndürme, "Sayfa n / N" yeniden numaralandırma | yalnız ilgili sayfa |

Korumalar (yardımcılar zorunlu kılar, agent atlayamaz):
- **Belirsiz hedef** (`Ambiguous`): ör. iki ayrı "Sonuç" başlığı → dosya
  yazılmaz, adaylar başlık yoluyla raporlanır.
- **Yıkıcı işlem** (`NeedsConfirmation`): formülün üzerine yazma, başka
  hücrelerin okuduğu satırı silme, başka slaytın köprü verdiği slaytı silme,
  formülü değere çevirme → yalnızca istekte `--onay` varsa. Bunu kütüphane
  de denetler: entrypoint istekteki `--onay`'ı ajan başlamadan `DOCGEN_ONAY`
  olarak konteynerin ana sürecine (PID 1) sabitler; ajanın `force=True`
  yazması ya da değişkeni kendisi ataması onay sayılmaz.
- **Yanlış "eski değer"** (`ValueMismatch`): dosyadaki gerçek değer raporlanır.
- **Yapılamayan** (`NotSupported`): PDF metninin rengi/fontu, sığmayan PDF
  metni, birleşik hücreli tabloda sütun ekleme, pivot tablolu sayfada satır
  işlemi vb. — daha geniş bir değişiklikle ikame edilmez.
- **"Hiçbir şeye dokunma"** isteği: dosya bayt bayt aynı teslim edilir.

---

## 4. Skill sistemi

Ajanın format bilgisi `/opt/docgen/skills/` altındaki dört skill'de
(SKILL.md = akış + kurallar, REFERENCE.md = API referans + OOXML anatomisi)
ve `manifest.json`'da tanımlıdır. Konteyner başlarken skill'ler
`/home/agent/.qwen/skills/` altına kopyalanır; QWEN.md "önce skill'i oku"
kuralını zorunlu kılar.

| Skill | Kapsam | Öne çıkanlar |
|---|---|---|
| `docx` | DOCX üret/düzenle/şablon | hızlı yol (markdown→docx), tam kontrol (docx-js), `docgen.docx` ile yerinde düzenleme (kapsam kuralları: ifade/paragraf/stil), kullanıcı şablonu (.dotx) akışı, hata modu tablosu |
| `pptx` | PPTX üret/düzenle/şablon | tema akışı (şablon yoksa), anti-AI tasarım kuralları, `build_themed_deck.js`, pptxgenjs, `docgen.pptx` ile yerinde düzenleme (konum/boyut, notlar, slayt işlemleri), kullanıcı şablonu (.potx) akışı |
| `xlsx` | XLSX üret/düzenle/şablon | formül kuralları, finans renk konvansiyonu (yeni dosyalar), `docgen.xlsx` ile yerinde düzenleme, `xlsx_struct.py` yapısal işlemler, `xlsx_recalc.py` zorunlu yeniden hesap, .xltx doldurma |
| `pdf` | PDF üret/işle/düzenle | HTML-öncelikli üretim, `pdf_text.py` / `pdf_pages.py` ile yerinde düzenleme, doğrulanmış Gotenberg rotaları, dürüst ret kuralları, taramalı PDF notu |

Her skill `pack_skill.py` ile taşınabilir `.skill` (zip) paketine çevrilebilir
(`skills/*.skill` — 4 paket üretildi).

---

## 5. Doğrulama araç seti (scripts/)

### 5.1 validate.py — yapısal + şema doğrulama
- Her zaman: zip bütünlüğü (`testzip`), `[Content_Types].xml` varlığı,
  tüm `.xml`/`.rels` parçalarının iyi biçimlenmiş XML olması.
- Şemalar mevcutken: ilgili parçaların ECMA-376 XSD doğrulaması
  (docx→`wml/wml__ROOT.xsd`, xlsx→`sml/sml_root.xsd`, pptx→`pml/pml-2ed.xsd`).
- PDF: `%PDF` başlığı + sayfa sayısı.
- Çıktı: insan-okunur veya `--json`; exit 0=geçti, 1=başarısız, 2=kullanım hatası.
- Yapısal hata varsa XSD adımı atlanır (sahte geçiş engellenir).
- `mc:Ignorable` uzantıları (w14/w15…) ve `mc:AlternateContent` doğrulamadan
  önce standart geri dönüşlerine indirgenir.
- **`--baseline <orijinal>`**: düzenlenen ya da şablondan üretilen dosyada
  yalnızca **yeni** şema bulguları hata sayılır. (Word'ün kendi yazdığı bazı
  yapılar — ör. `w:percent`'siz `w:zoom` — 1. baskı şemasına uymaz; bunlar
  "pre-existing" olarak listelenir.) `.dotx/.potx/.xltx` ve makrolu uzantılar
  da kabul edilir.

### 5.2 diff_check.py — "yalnızca istenen değişti mi?" (düzenlemelerde zorunlu)
Orijinal ile sonucu içerik düzeyinde karşılaştırır; ajanın bildirdiği
kapsam dışındaki her değişiklikte FAIL verir.
- DOCX: paragraf metni, karakter düzeyinde biçim (biçim "taşması" dahil),
  paragraf biçimi, resimler, stiller, tablolar, sayfa yapısı.
- PPTX: şekil metni/biçimi/konum-boyutu/dolgusu, notlar, arka plan, gizlilik,
  slayt sırası, eklenen/silinen slaytlar.
- XLSX: değer/formül, hesaplanmış sonuç (sayısal tolerans), görsel olarak
  çözümlenmiş hücre biçimi (tema rengi/fontu değerlere çevrilir), koşullu
  biçim, doğrulama, birleşik hücreler, sütun genişlikleri, tablolar, adlar,
  grafikler, çizimler, dış bağlantılar.
- PDF: sayfa eşleme (silinen/eklenen/taşınan), satır metni, font, punto,
  renk, konum; üst/alt bilgi bandı ayrı.
- Paket düzeyi: kapsam dışı parça değişikliği ve ZIP klasör kayıtları.
- Kapsam: `--allow-text/-section/-style/-slide/-notes/-page/-band/-cell/
  -sheet/-recalc/-feature/-structure/-theme/-part`, `--text-only`,
  `--format-only`; beklenti: `--expect-text`, `--expect-gone`;
  `--record outputs/manifest.json` (orijinal/sonuç SHA-256 + değişiklikler).

### 5.3 find_target.py — hedef bulma
Bir metnin tüm geçişlerini bağlamıyla listeler (DOCX başlık yolu ve tablo
hücresi, PPTX slayt/şekil, XLSX hücre, PDF sayfa/satır). Exit 0 = tek
eşleşme, 3 = birden fazla (belirsiz), 1 = yok.

### 5.4 readback.py — içerik geri okuma
Üretilen dosyayı office_oxide (ofis) veya pdfplumber (PDF) ile geri okur.
`--format`: run/paragraf biçimi, şekil konumu, hücre biçimi, PDF satır
font/punto/renk/konumu — biçim değişikliklerini okuyarak doğrulamak için.
`--tables`, `--max-chars`; `.dotx/.potx/.xltx` da okunur.

### 5.5 render_check.py — render kontrolü
Ofis dosyasını Gotenberg'in LibreOffice'una açtırıp PDF'e döktürür; gerçekten
sayfa üretmediyse FAIL verir. Elle yazılan `curl -s … -o x.pdf && ls x.pdf`
kalıbının yerini alır: `-f` olmadan curl, Gotenberg'in 4xx/5xx hata metnini
"x.pdf" olarak kaydedip exit 0 döner, başarısız render başarılı görünür.
- Gotenberg 200 dönmeli, gövde açılan bir PDF olmalı, en az 1 sayfa.
- Sunumlarda PDF sayfa sayısı = görünür slayt sayısı (gizli slaytlar basılmaz).
- Dosyada metin var ama PDF'te yoksa FAIL (boş render).
- `--expect-pages N`, `--expect-text T` (tekrarlanabilir), `-o` ile PDF'i
  saklama, `--json`; Gotenberg meşgulse (502/503/504, bağlantı hatası) bir
  kez yeniden dener.
- Formül hataları burada aranmaz: `xlsx_recalc.py` hesaplanan değerleri
  doğrudan okur ve hangi hatanın yeni olduğunu bilir.
- Eval'de docx/pptx/xlsx teslim eden her vakada zorunlu adım (`render`).

### 5.6 xlsx_recalc.py — formül sonuçları
LibreOffice'i yalnızca hesap motoru olarak kullanır: kopyayı hesaplatır,
sonuçları **orijinal paketin** formül hücrelerine önbellek değeri olarak
yazar (başka hiçbir parça değişmez). Yeni `#REF!/#DIV/0!…` hatalarında ve
`--expect "Sayfa!B4=20000"` tutmadığında exit 1. (`recalc_check.py`,
LibreOffice yoksa Gotenberg üzerinden tarama yapan eski yedek yoldur.)

### 5.7 xlsx_struct.py — satır/sütun ve sıralama
`insert-rows/delete-rows/insert-cols/delete-cols/sort` sayfa XML'inde yerinde
yapılır; sayfaya giden tüm referanslar Excel kuralıyla kaydırılır (tüm
sayfalardaki formüller, adlar/yazdırma başlıkları, CF/DV aralık ve
formülleri, tablolar, birleşik hücreler, köprüler, yorumlar, çizim
çapaları, grafik serileri). Okunan satırı silme `--onay` ister; ardından
`xlsx_recalc.py` otomatik çalışır.

### 5.8 pdf_text.py / pdf_insert.py / pdf_pages.py — PDF düzenleme
`pdf_text.py`: metni yerinde değiştirir — yeni karakterlerin hepsi gömülü
fontta varsa ve satır sığıyorsa mevcut metin nesnesi; aksi halde paragraf
aynı font ailesinin tam dosyasıyla sütun içinde yeniden çizilir. Karışık
biçimli satır, tablo satırının tamamı, kurulu olmayan font ailesi veya
sığmayan metinde reddeder; sonucu sayfayı geri okuyarak doğrular.
`pdf_insert.py`: verilen satırın altına yeni paragraflar ekler — sayfanın
gövde fontu, punto, renk, satır aralığı ve sütunuyla; sayfada başka hiçbir
şey kaymaz. Yeniden akış (reflow) yoktur: metin sonraki bloğa/alt bilgiye
kadar sığmıyorsa reddeder (gereken/kalan satır sayısını söyler).
`pdf_pages.py`: delete/keep/move/insert/rotate + `--renumber`.

### 5.9 convert_legacy.py
`.doc/.xls/.ppt` (ve `.rtf/.odt/.ods/.odp`) → OOXML; dönüşüm son yanıtta
belirtilir.

### 5.10 recalc_check.py — eski formül QA'sı (yedek)
Dosyayı Gotenberg'in LibreOffice'u üzerinden PDF'e çevirip hata
tokenlarını tarar; değer yazamaz. LibreOffice'in agent imajında olmadığı
durumlar için tutulur.

### 5.11 visual_qa.py — görsel QA eklentisi (varsayılan: KAPALI)
PDF'i sayfa sayfa PNG'ye render eder; `--contact-sheet` ile tek ızgara
görüntüsü, `--docling` ile Docling kuruluysa OCR çapraz kontrolü
(bos/bozuk sayfa tespiti). İstekte `--visual-qa` gelmediyse çalıştırılmaz —
Gemma'nın vision yeteneği olmadığı için normal akışa girmez.

### 5.12 Yardımcılar
- `repair_pptx.py` — pptxgenjs çıktısındaki bilinen iki XML kusurunu
  düzeltir: `presentation.xml` öğe sırası (XSD'ye göre
  `notesMasterIdLst` → `sldIdLst` sıralaması) ve placeholder'da birleşik
  paragraflar (aynı `<a:p>` içindeki tekrarlı `<a:pPr>` → ayrı paragraflar).
  Kendiliğinden (idempotent): düzeltecek şey yoksa "clean" der. Her
  pptxgenjs build'i sonrası, validate'den ÖNCE çalışır.
- `build_templates.py` — imaj derleme sırasında çalışır:
  `themes/templates/docx/base.docx` (placeholder'lı iskelet) ve
  `themes/templates/pptx/basic.pptx` (10 slayt) üretir.
- `patch_qwen_prompt.py` — imaj derleme sırasında çalışır: Qwen Code'un
  headless sistem talimatındaki "tek turluk çalışma, asla soru sorma,
  varsayıp bitir" cümlesini docgen'in gerçeğiyle değiştirir (bkz. Ajan
  çekirdeği). Cümle tam olarak bir kez bulunmazsa derleme durur; Qwen
  sürümü yükseltilince betikteki metin gözden geçirilir.
- `build_themed_deck.js` — standart temalı sunum üretici
  (tema JSON + içerik JSON → pptxgenjs ile kurallara uygun deck).
- `pack_skill.py` — skill dizinini `.skill` paketini çevirir.
- `e2e_engine_test.sh` — model bağımsız uçtan uca motor testi
  (xlsx + pdf zincirleri; §9'da gösterildiği gibi konteynerde çalışır).

---

## 6. Tema sistemi

`/opt/docgen/themes/` altında üç hazır tema (JSON token dosyası):

| Tema | primary | Amaç |
|---|---|---|
| `professional-blue` (varsayılan) | #1F4E79 | kurumsal/varsayılan |
| `minimal` | #111111 | sade, tek ton |
| `dark-corporate` | #4FC3F7 | koyu kurumsal (ekran; baskıda CSS beyaz zemin kullanır) |

Token şeması: `colors{primary,secondary,accent,background,text,muted}`,
`fonts{heading,body}`, `spacing`, `rules{max_bullets_per_slide, max_fonts,
accent_lines_under_titles}`.

- HTML→PDF için her temanın yazdırılabilir CSS preset'i var
  (`themes/css/*.css`: A4 @page, tablo başlığı tekrarı, page-break kuralları).
- PPTX için `build_themed_deck.js` token'ları master'a uygular.
- **Kullanıcı şablonu her docgen temasını ezer:** `.dotx/.docx`,
  `.potx/.pptx` veya `.xltx/.xlsx` şablon olarak verilirse stilleri,
  master/düzenleri, teması, üst/alt bilgisi (logo) korunur ve içerik bu
  şablon üzerine kurulur (`start_from_template` + python-docx/python-pptx;
  xlsx için şablon kopyası yerinde doldurulur). Doğrulama `--baseline <şablon>`.
- Agent kuralı: tema token'ları **varsayılandır**; agent kendiliğinden tema
  dışı renk/font üretmez. Tema seçimi final raporda belirtilir.
- **Kullanıcının açık isteği temayı ezer:** "başlığı kırmızı yap", "bu
  tablonun başlıklarını koyu yeşil yap", "font 14 punto olsun" gibi
  isteklerde tam olarak istenen uygulanır, tema dışında olsa bile, ve
  yalnızca istenen kapsamda (tek başlık ↔ tüm başlıklar). Bu kural skill'lerdeki
  tema/font/renk konvansiyonu kurallarından da önce gelir (bkz. Bölüm 3.1).

**Anti-AI stil kuralları (tüm formatlar):** başlık altı dekoratif çizgi/çubuk
yok; slayt başına en fazla 6 kısa madde; en fazla 2 font ailesi; her
slayt/bölümde tutarlı düzen; dolgu metin/uysal içerik yok.

---

## 7. XSD şema seti (ECMA-376)

`/opt/docgen/schemas/` altında 178 XSD dosyası — OOXML (Office Open XML)
şema setinin docx4j aynasından alınmış kopyası (Apache-2.0).

- **Önemli:** dizin yapısı KESİNLİKLE DÜZGÜNE ÇEKİLMEZ — XSD'ler aralarında
  göreli import kullanır (`../dml/`, `../shared/` vb.).
- Giriş şemaları: `wml/wml__ROOT.xsd` (docx), `sml/sml_root.xsd` (xlsx),
  `pml/pml-2ed.xsd` (pptx).
- Şemalar yoksa / `DOCGEN_SCHEMAS` boşsa validate.py yapısal kontrolle
  sınırlı kalır ve uyarı verir; doğrulama zinciri kırılmaz.

---

## 8. Proje dosya ağacı

```
docgen/
├── Dockerfile                  # agent imajı (python 3.12 + node 22 + qwen + motorlar)
├── docker-compose.yml          # gotenberg + agent servisleri (iç ağ)
├── .env.example                # MODEL_ENDPOINT / MODEL_NAME / MODEL_API_KEY
├── .env                        # (gitignore) kendi model yapılandırmanız
├── requirements.txt            # Python bağımlılıkları (lisanslı)
├── package.json                # Node bağımlılıkları: docx, pptxgenjs
├── QWEN.md                     # ajan sistem kuralları (dosya modeli, skill zorunluluğu,
│                               #   tema/anti-AI kuralları, doğrulama döngüsü, Gotenberg)
├── entrypoint.sh               # skill kopyalama + Gotenberg health + qwen headless
├── data/                       # (bind mount -> /workspace)
│   ├── uploads/                # kullanıcı kaynakları (:ro mount)
│   ├── scratch/                # çalışma alanı
│   ├── outputs/                # TESLİMATLAR (+ manifest.json)
│   └── ui/                     # sohbet arayüzü oturumları (Bölüm 13)
├── ui/                         # sohbet arayüzü (host'ta: python3 ui/server.py)
│   ├── server.py               # stdlib HTTP sunucu: oturumlar, agent konteynerleri, önizleme
│   ├── xlsx_view.py            # Excel önizleme modeli: biçimli değerler, formüller, stiller
│   └── static/                 # index.html, app.js, sheet.js, style.css (derleme adımı yok)
├── lib/docgen/                 # agent'ın kendi koduyla kullandığı kütüphane
│   ├── package.py              # OOXML paketi: yalnız değişen parça yazılır
│   ├── docx.py, pptx.py, xlsx.py, xlsx_struct.py, pdf.py, theme.py
│   ├── compare.py              # diff_check modelleri
│   └── office.py               # LibreOffice (hesap/dönüşüm)
├── scripts/
│   ├── validate.py             # yapısal + XSD doğrulama (--baseline)
│   ├── diff_check.py           # kapsam dışı değişiklik kontrolü
│   ├── find_target.py          # hedef bulma (belirsizlik tespiti)
│   ├── readback.py             # içerik / biçim geri okuma
│   ├── render_check.py         # LibreOffice (Gotenberg) render kontrolü
│   ├── xlsx_recalc.py          # formül sonuçlarını hesaplayıp yazar
│   ├── xlsx_struct.py          # satır/sütun/sıralama (referans kaydırma)
│   ├── pdf_text.py, pdf_insert.py, pdf_pages.py   # PDF metin / paragraf / sayfa düzenleme
│   ├── convert_legacy.py       # .doc/.xls/.ppt -> OOXML
│   ├── recalc_check.py         # eski formül QA'sı (Gotenberg yedeği)
│   ├── visual_qa.py            # görsel QA eklentisi (varsayılan kapalı)
│   ├── repair_pptx.py          # pptxgenjs XML kusurlarını düzeltir (idempotent)
│   ├── build_templates.py      # imaj derleme: docx/pptx baz şablonları
│   ├── patch_qwen_prompt.py    # imaj derleme: Qwen headless "soru sorma" cümlesini değiştirir
│   ├── build_themed_deck.js    # standart temalı sunum üretici
│   ├── e2e_engine_test.sh      # model bağımsız uçtan uca motor testi
│   └── pack_skill.py           # skill -> .skill paketi
├── skills/
│   ├── manifest.json
│   ├── docx/{SKILL.md,REFERENCE.md}
│   ├── pptx/{SKILL.md,REFERENCE.md}
│   ├── xlsx/{SKILL.md,REFERENCE.md}
│   ├── pdf/{SKILL.md,REFERENCE.md}
│   └── {docx,pptx,xlsx,pdf}.skill   # pack_skill.py ile üretilen paketler
├── themes/
│   ├── professional-blue.json  # varsayılan
│   ├── minimal.json
│   ├── dark-corporate.json
│   ├── css/{professional-blue,minimal,dark-corporate}.css
│   └── templates/{docx/base.docx, pptx/basic.pptx}   # imaj derlemede üretilir
└── schemas/                    # 178 XSD (ECMA-376) + README.md (kaynak notu)
```

---

## 9. Hızlı başlangıç

**Gereksinimler:** Docker (v20+), `gotenberg/gotenberg:8` imajı (yerelde
yoksa ilk build'de çekilir), kendi OpenAI-uyumlu model endpoint'iniz.

```bash
# 1. yapılandırma
cp .env.example .env
#    .env içine:
#      MODEL_ENDPOINT=http://<host-makineniz>:8080/v1   # OpenAI-uyumlu /v1
#      MODEL_NAME=gemma-3-27b-it                         # servisinizin model adı
#      MODEL_API_KEY=                                    # gerekiyorsa

# 2. derle + çalıştır (tek istek, tek konteyner)
docker compose build
docker compose run --rm agent "Satış ekibi için 2026 Q3 özeti, 5 slaytlık temalı bir .pptx hazırla"

# kaynak dosya içeren istek: dosyayı data/uploads/ koyun
docker compose run --rm agent "uploads/rapor.docx içindeki rakamlarla 3 sayılı bir xlsx özet tablosu oluştur"
```

Teslimatlar `data/outputs/` altında olur. Her istek için konteyner
`--rm` ile sökülür; Gotenberg servisi arka planda kalır (restart
`unless-stopped`).

**Çalışma örnekleri (model bağımsız test):** motor ve QA zinciri model
endpoint'i gerektirmez — hazır bir uçtan uca motor testi var
(xlsx: oluştur→validate→readback→recalc→render_check, bozuk dosyada render_check
FAIL vermeli; pdf: html→chromium→validate→readback):

```bash
docker compose run --rm --entrypoint bash agent -c 'bash /opt/docgen/scripts/e2e_engine_test.sh'
```

Tekil parça denemek için de elle istek verilebilir:

```bash
docker compose run --rm --entrypoint bash agent -c '
  python /opt/docgen/scripts/build_templates.py && \
  python /opt/docgen/scripts/validate.py /opt/docgen/themes/templates/docx/base.docx && \
  python /opt/docgen/scripts/readback.py /opt/docgen/themes/templates/pptx/basic.pptx | head -5'
```

---

## 10. Genişletme noktaları

| Genişletme | Nasıl |
|---|---|
| **Görsel QA (vision)** | `--visual-qa` isteğinde `visual_qa.py` PNG/contact-sheet üretir; Docling kurulu ortamda `--docling` ile OCR çapraz kontrolü eklenir. Normal akışa dokunmaz. |
| **Yeni tema** | `themes/<ad>.json` (token şeması) + `themes/css/<ad>.css` ekleyin; manifest/skill dokümanına ekleyin. |
| **Yeni format skill** | `skills/<ad>/{SKILL.md,REFERENCE.md}` + `manifest.json` kaydı; `pack_skill.py` ile paketleyin. |
| **Farklı LLM** | `.env`'i değiştirin — OpenAI-uyumlu herhangi bir endpoint/model adı çalışır (model seçimi mimariye dokunmaz). |
| **Yeni QA adımı** | `scripts/` altına ek script + QWEN.md doğrulama döngüsüne ekleme. |

---

## 11. Güvenlik modeli

- Ajan **non-root** (`agent`, uid 1000) kullanıcı; bağlı `data/` dışında
  hosta yazamaz.
- `/opt/docgen/` root'a ait, agent için salt okunur (dosya izinleriyle).
- `uploads/` compose'da ayrıca `:ro` bağlanır: orijinal dosya mount
  düzeyinde korunur (yalnızca talimata dayanmaz).
- Konteyner **tek seferlik** (`--rm`): oturum state'i birikmez, her istek
  temiz sandbox'ta başlar.
- Gotenberg host portu **açılmaz**; yalnızca compose iç ağından erişilir.
- Model anahtarı yalnızca `.env` (gitignore) içinde; konteynere env olarak
  enjekte edilir, hiçbir dosyaya yazılmaz.
- XML ayrıştırma: docgen kütüphanesi lxml'i `resolve_entities=False,
  no_network=True` ile kullanır; `validate.py`/`repair_pptx.py` lxml'in
  varsayılanlarıyla çalışır (lxml ≥ 5 dış entity çözmez). defusedxml kurulu
  ama şu an hiçbir script tarafından kullanılmıyor. Ofis dosyaları
  zip-bütünlüğü kontrolünden geçer.
- Açık kalanlar (bilinçli): agent'ın internet erişimi vardır (model
  servisine ulaşması gerekir) ve `--yolo` ile komutları onaysız çalıştırır;
  `MODEL_API_KEY` ortam değişkeni olarak konteynerde görünürdür; CPU/bellek
  limiti tanımlı değildir.

---

## 12. Yetenek eval'i (`evals/`)

Gerçek sistemi (aynı imaj, aynı entrypoint, `docgen` ağı + Gotenberg, `.env`
ile Qwen agent) dört format için uçtan uca ölçer.

| Dosya | Görev |
|---|---|
| `evals/cases.csv` | 234 vaka (DOCX 78, PPTX 62, XLSX 47, PDF 47): istek, beklenen sonuç türü (`edit/create/refuse/partial/nochange`), bağımsız kontroller, referans çözüm, zorunlu/yasak araç adımları, hakem ölçütü |
| `evals/build_cases.py` | vakaları fixture içeriğinden üretir; hedef metinlerin var/benzersiz olduğunu doğrular |
| `evals/fixtures/` | test belgeleri (+ `.doc/.ppt/.xls` eski format kopyaları) |
| `evals/checks.py` | **bağımsız** kontroller: python-docx / python-pptx / openpyxl / pdfplumber / ham OOXML; formül sonuçları taze LibreOffice hesabıyla karşılaştırılır — docgen kodu kullanılmaz |
| `evals/reference.py` | her vakanın referans çözümü (self-test için) |
| `evals/evaluate.py` | container içinde puanlama: sonuç türü, dosya kontrolleri, araç yolu (Qwen Code JSON izinden), Gemma hakem |
| `evals/run_eval.py` | host tarafı koşucu (yalnız stdlib + docker) |

Puanlama katmanları: **sonuç** (dosya üretildi mi / doğru şekilde üretilmedi
mi), **dosya kontrolleri**, **araç yolu** (ör. düzenlemede `find_target` →
`docgen.*` → `validate --baseline` → `diff_check` → `render_check`; `EditableDocument`,
`create_from_markdown`, `--onay` atlatması gibi yasak yollar), **hakem**
(Gemma, LiteLLM `small_llm`: yanıt dosyayla tutarlı ve dürüst mü). Vaka ancak
dördünden de geçerse başarılı sayılır; başarısızlığın nedeni (yol / sonuç /
yanıt / altyapı) raporlanır.

```bash
python3 evals/run_eval.py --selftest                 # önce: kontroller doğru mu? (referans çözümler geçmeli,
                                                     #   dokunulmamış dosya kalmalı)
python3 evals/run_eval.py --concurrency 6            # tüm vakalar -> evals/runs/<zaman>/report.md
python3 evals/run_eval.py --filter 'PPTX|XLSX-structure'
python3 evals/run_eval.py --grade-only evals/runs/<run>   # yeniden puanla
```
`.env`'de `MODEL_*` (agent: Qwen) ve `GEMMA_*` (hakem) tanımlı olmalı.
`build_cases.py` değiştirildiyse `python evals/build_cases.py` ile CSV yeniden
üretilir.

---

## 13. Sohbet arayüzü (`ui/`)

Tarayıcıdan belge isteyip sonucu yanda önizlemek ve yüklenen dosyada
değişiklik istemek için. Host'ta çalışır (`run_eval.py` gibi yalnız Python
stdlib + docker); her mesaj CLI ile aynı imaj, entrypoint, `docgen` ağı ve
`.env` ile tek seferlik bir agent konteyneri başlatır.

```bash
docker compose build          # entrypoint değiştiyse bir kez (canlı adımlar için)
python3 ui/server.py          # http://127.0.0.1:8765   (--port, --host, --timeout, --image)
```

| Özellik | Nasıl |
|---|---|
| **Sohbet** | Her mesaj ayrı bir konteynerde çalışır; son 6 mesaj ve yanıtları bağlam olarak isteğin başına eklenir. Bir sohbette aynı anda tek istek çalışır. |
| **Dosya yükleme** | Ataş düğmesi ya da sürükle-bırak: `.docx/.pptx/.xlsx/.pdf`, şablonlar (`.dotx/.potx/.xltx`), makrolu ve eski biçimler (`.docm/.pptm/.xlsm`, `.doc/.xls/.ppt/.rtf/.odt/.ods/.odp`); kaynak olarak `.csv/.txt/.md/.json/.png/.jpg` (dosya başına en fazla 200 MB). |
| **Düzenleme** | Önceki yanıtların çıktıları sonraki mesajın `uploads/`'ına (salt okunur) kopyalanır: "başlığı kırmızı yap" son çıktıyı yerinde düzenler. Önizlemede açık olan dosya, adı söylenmeyen düzenleme isteğinin hedefi olarak ajana bildirilir. |
| **Önceki çalışma dosyaları** | Her önceki mesajın `scratch/`'ı `/workspace/previous/<n>/` olarak salt okunur bağlanır ve prompt'ta hangi mesajın neyi teslim ettiğiyle listelenir. Böylece HTML'den üretilmiş bir PDF'te renk/tablo/bölüm gibi yerinde yapılamayan değişiklikler kaynağı düzenleyip yeniden render ederek yapılır (QWEN.md, `pdf` skill). Kaynak yalnız aynı mesajın teslimatına aittir. |
| **Zaman aşımı** | Mesaj başına varsayılan 1500 sn (`--timeout`). Zaman aşımına uğrayan, durdurulan ya da yarıda kalan mesaj sonraki isteklerin geçmişinde "tamamlanmadı" olarak geçer; yarım işi `/workspace/previous/<n>/` altında görünür ama doğrulanmamış sayılır. |
| **Önizleme** | Ofis dosyaları Gotenberg (LibreOffice) ile PDF'e çevrilip önbelleğe alınır ve sağ panelde gösterilir; PDF, görsel ve metin doğrudan gösterilir. "İndir" her zaman asıl dosyayı verir. |
| **Excel önizleme** | `.xlsx/.xlsm/.xltx/.csv` varsayılan olarak Excel benzeri tabloda açılır (`ui/xlsx_view.py` → `ui/static/sheet.js`; "Tablo / PDF" ile geçilir): sayfa sekmeleri, formül çubuğu (hücreye tıklayınca formül ve hesaplanmış sonuç), "Formüller" düğmesi (Ctrl+\`) ile tüm formüller hücrelerde, ok tuşlarıyla gezinme. Değerler Excel'in sayı/tarih biçimiyle ve Türkçe ayraçlarla gösterilir (sığmayan sayı `####`); paylaşılan formüller her hücre için açılır. Stil: yazı tipi, dolgu, kenarlık, hizalama, tema rengi/tint, birleştirme, sütun genişliği/satır yüksekliği, gizli satır/sütun, dondurulmuş bölmeler, hücre açıklaması, veri doğrulama bilgisi, `cellIs`/metin/renk skalası koşullu biçimleri ve yerleşik tablo stillerinin yaklaşık hali. Grafik, görsel, veri çubuğu/simge kümesi ve formüllü koşullu biçimler tabloda yoktur; sayfada not olarak belirtilir, PDF görünümünde görülür. |
| **Canlı adımlar** | Arayüz `DOCGEN_OUTPUT_FORMAT=stream-json` ve `DOCGEN_OUT_FILE=/workspace/.docgen-events.jsonl` verir; entrypoint olayları bu dosyaya yazar, arayüz onları adım listesi (komut, okunan/yazılan dosya, çıktı) olarak gösterir. Bu değişkenleri tanımayan eski imajda adımlar iş bitince görünür. |
| **`--onay`** | Mesaj kutusundaki "onay" düğmesi isteğe `--onay` ekler; yalnız o mesaj için geçerlidir (önceki mesajlardaki `--onay` bağlama taşınmaz). |
| **Durdur** | Çalışan konteyneri `docker kill` ile durdurur; o ana kadar üretilen çıktılar yine toplanır. |

Oturumlar `data/ui/<oturum>/` altında tutulur (gitignore): `files/` (tüm
yüklemeler ve çıktılar), `turns/<n>/` (o mesajın `/workspace`'i, `prompt.txt`,
ham ajan çıktısı) ve `previews/` (PDF önizleme önbelleği).

**Güvenlik:** sunucu varsayılan olarak yalnız `127.0.0.1`'e bağlanır ve
kimlik doğrulaması yoktur — `--host 0.0.0.0` ile açılırsa ağdaki herkes ajanı
çalıştırabilir ve oturum dosyalarını indirebilir. Ajan tarafındaki sandbox
modeli (non-root, `uploads/` salt okunur, `--rm`) CLI ile aynıdır.

---

## 14. MCP servisi (`src/docgen_service/`, `deploy/`)

Doküman ajanını bir host sisteme (Oppy) MCP üzerinden bağlayan ince katman.
Yapısı oppy-web-agent'ın MCP katmanıyla aynıdır; entegrasyon rehberi
[docs/INTEGRATION.md](docs/INTEGRATION.md), planlayıcı için yetenek tanımı
[docs/capability/document_capability.yaml](docs/capability/document_capability.yaml),
yönlendirme test seti
[docs/capability/routing_test_set.yaml](docs/capability/routing_test_set.yaml).

| Parça | Ne yapar |
|---|---|
| `interfaces/mcp_server.py` | İnce adaptör: `document_run`, `document_job`, `document_cancel`, `document_read`, `document_status`; `docgen://status` ve `docgen://jobs/{id}/files/{ad}` resource'ları; `/health` (token yok), `/jobs/{id}/files/{ad}` (token); bearer token, eşzamanlılık kapıları, SIGTERM'de boşaltma |
| `jobs.py` | İş servisi: her iş kendi `docgen-agent` konteynerinde, UI ile aynı şekilde (aynı entrypoint, ağ, `.env`, stream-json); kuyruk, zaman aşımı, iptal, `continue_from` zinciri (dosyalar + `/workspace/previous/<n>/`), kalıcı kayıt, saklama süresi |
| `conversation.py` | UI ile paylaşılan kısım: prompt kurma, olay → adım özeti, dosya adları (`ui/server.py` de bunu kullanır) |
| `core/availability.py` | Öz-denetim: docker, ajan imajı, Gotenberg, model endpoint'i, ajan `.env`'i |
| `config.py` | Tüm ayarlar `DOCGEN_` önekiyle; ajanın `.env`'inden ayrı (token ajana asla geçmez) |

```bash
./deploy/generate-secrets.sh                      # deploy/stack.env: token, iş klasörü
cp deploy/service.env.example deploy/service.env
docker compose -f deploy/compose.yml --env-file deploy/stack.env build
docker compose -f deploy/compose.yml --env-file deploy/stack.env up -d docgen-mcp
curl http://127.0.0.1:18766/health
python deploy/check_mcp.py --token-file deploy/stack.env   # uçtan uca kontrol (gerçek iş)
.venv/bin/python -m pytest                                 # birim + sözleşme + drift testleri
```

Onay (`--onay`) MCP'de yalnızca `confirm` parametresiyle verilir: istek
metnindeki `--onay` yok sayılır (planlayıcı modelin yazdığı metin kullanıcı
onayı değildir). Servis konteyneri işleri hostun docker soketiyle başlatır —
bu, hostta root yetkisi demektir; port yalnızca loopback'te yayınlanır
(bkz. INTEGRATION.md §1).
