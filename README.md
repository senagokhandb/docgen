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
| **gotenberg** | `gotenberg/gotenberg:8` | Render ve PDF motoru: Chromium (HTML→PDF), LibreOffice (ofis→PDF), pdfengines (merge/crop). Host portu açılmaz; yalnızca iç ağdan erişilir. |

### 1.2 Agent konteynerinin içindekileri

| Katman | İçerik |
|---|---|
| **Ajan çekirdeği** | Qwen Code CLI v0.26.1 — istemi headless modda (`--yolo -o json`) çalıştırır; kod yazma, shell komutu, dosya okuma/yazma yetkileri tam |
| **Model** | `.env` dosyasındaki OpenAI-uyumlu endpoint (Gemma). Konteyner dışı, kendi inference servisiniz |
| **Python motoru** | office_oxide (ofis okuma/yazma/düzenleme), pdfplumber, lxml, openpyxl, requests, defusedxml, pypdfium2, Pillow |
| **Node motoru** | docx-js (docx), pptxgenjs — `NODE_PATH=/opt/docgen/node_modules` |
| **Bilgi katmanı** | `/opt/docgen/skills/` (4 format skill + referanslar), `/opt/docgen/QWEN.md` (sistem kuralları), `/opt/docgen/themes/` (temalar + CSS + şablonlar), `/opt/docgen/schemas/` (ECMA-376 XSD seti) |
| **Doğrulama araç seti** | `/opt/docgen/scripts/` — validate, readback, recalc_check, visual_qa, repair_pptx, build_templates, build_themed_deck, pack_skill |

### 1.3 Dosya modeli (sandbox kuralları)

| Yol | Erişim |
|---|---|
| `/workspace/uploads/` (host: `data/uploads/`) | salt okunur — kullanıcı kaynakları |
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
   b. Tema seçer (kullanıcı belirtmediyse professional-blue, bunu bildirir)
   c. Motoru seçer (skill'deki tabloya göre) ve kodu /workspace/scratch/ yazar
   d. Dosyayı üretir
   e. DOĞRULAMA DÖNGÜSÜ:
        validate.py  -> yapısal + ECMA-376 XSD
        readback.py  -> üretilen dosyanın içeriğini geri oku
        Gotenberg    -> PDF render (sayfa/slayt sayısı kontrolü)
        recalc_check.py -> YALNIZCA formüllü xlsx (yeniden hesap + beklenen değerler)
      Başarısız adım: düzelt + tekrar (maks. 3 tur, sonra dürüst rapor)
   f. Son dosyayı /workspace/outputs/ koyar; yol + 2-3 cümle özetle raporlar
```

Görsel QA **varsayılan olarak kapalıdır**; istekte `--visual-qa` varsa
`visual_qa.py` devreye girer (Bölüm 5.4).

---

## 3. Motor matrisi

| İhtiyaç | Araç | Dil |
|---|---|---|
| Ofis dosyası okuma (markdown/HTML/düz metin) | office_oxide `Document` | Python |
| Markdown'dan DOCX/PPTX/XLSX (hızlı yol) | office_oxide `create_from_markdown` | Python |
| XLSX üretimi: hücre, formül, stil, birleştirme | office_oxide `XlsxWriter` | Python |
| PPTX üretimi (basit/sadece metin) | office_oxide `PptxWriter` | Python |
| Mevcut ofis dosyasında metin/hücre düzenleme | office_oxide `EditableDocument` | Python |
| Hassas xlsx hücre/stil/formül (mevcut dosya) | openpyxl | Python |
| DOCX tam kontrol (tarz, başlık/altlık, tablo) | docx (docx-js) | Node |
| PPTX tam kontrol (şema, grafik, görsel) | pptxgenjs | Node |
| PDF metin/tablo çıkarma (QA geri-okuma) | pdfplumber | Python |
| XML parse + XSD doğrulama | lxml | Python |
| Güvenli XML parse | defusedxml | Python |
| PDF→PNG render (görsel QA) | pypdfium2 + Pillow | Python |
| HTTP istemcisi (Gotenberg) | requests | Python |
| HTML→PDF | Gotenberg Chromium | servis |
| Ofis→PDF, xlsx yeniden hesap | Gotenberg LibreOffice | servis |
| PDF merge/crop/sayfa seçimi | Gotenberg pdfengines | servis |

Tüm bağımlılıklar açık kaynak (MIT / BSD / Apache-2.0 / HPND / PSF);
lisanslar `requirements.txt` ve `package.json` içinde satır satır belirtilir.

---

## 4. Skill sistemi

Ajanın format bilgisi `/opt/docgen/skills/` altındaki dört skill'de
(SKILL.md = akış + kurallar, REFERENCE.md = API referans + OOXML anatomisi)
ve `manifest.json`'da tanımlıdır. Konteyner başlarken skill'ler
`/home/agent/.qwen/skills/` altına kopyalanır; QWEN.md "önce skill'i oku"
kuralını zorunlu kılar.

| Skill | Kapsam | Öne çıkanlar |
|---|---|---|
| `docx` | DOCX üret/düzenle | hızlı yol (markdown→docx), tam kontrol (docx-js), unzip→lxml→rezip düzenleme, hata modu tablosu |
| `pptx` | PPTX üret/düzenle | zorunlu tema akışı, anti-AI tasarım kuralları, `build_themed_deck.js` hızlı üretici, pptxgenjs tam kontrol, "yeniden üret > yama" kuralı |
| `xlsx` | XLSX üret/düzenle | formül kuralları (desteklenmeyen işlevler + `_xlfn.` listesi), finans renk konvansiyonu, XlsxWriter API tuzakları (0 tabanlı satır/sütun, formülde "="), yeniden hesap QA'sı |
| `pdf` | PDF üret/işle | HTML-öncelikli üretim (temalı CSS preset), pdfengines işlemleri, "kaynak dosyayı bul" kuralı, taramalı PDF→Docling notu |

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

### 5.2 readback.py — içerik geri okuma
Üretilen dosyayı office_oxide (ofis) veya pdfplumber (PDF) ile geri okur;
ajan "niyet ettiğim içerik gerçekten dosyada mı?" sorusunu yanıtlar.
`--tables` (xlsx/pdf tablo özeti), `--max-chars`.

### 5.3 recalc_check.py — formül QA'sı (xlsx)
Dosyayı Gotenberg'in LibreOffice'u üzerinden PDF'e çevirir (LibreOffice
formülleri yüklemeyi yeniden hesaplar), pdfplumber ile metni tarar:
- Hata tokenları: `#REF! #NAME? #DIV/0! #VALUE! #N/A #NULL! #NUM!`
- `--expect "Etiket:değer"` (tekrarlanabilir): beklenen hesaplanan değerler
  render edilmiş sayfada aranır (boşluk/virgül normalizasyonu ile).
- Exit 0/1/2 + `--json` rapor.

### 5.4 visual_qa.py — görsel QA eklentisi (varsayılan: KAPALI)
PDF'i sayfa sayfa PNG'ye render eder; `--contact-sheet` ile tek ızgara
görüntüsü, `--docling` ile Docling kuruluysa OCR çapraz kontrolü
(bos/bozuk sayfa tespiti). İstekte `--visual-qa` gelmediyse çalıştırılmaz —
Gemma'nın vision yeteneği olmadığı için normal akışa girmez.

### 5.5 Yardımcılar
- `repair_pptx.py` — pptxgenjs çıktısındaki bilinen iki XML kusurunu
  düzeltir: `presentation.xml` öğe sırası (XSD'ye göre
  `notesMasterIdLst` → `sldIdLst` sıralaması) ve placeholder'da birleşik
  paragraflar (aynı `<a:p>` içindeki tekrarlı `<a:pPr>` → ayrı paragraflar).
  Kendiliğinden (idempotent): düzeltecek şey yoksa "clean" der. Her
  pptxgenjs build'i sonrası, validate'den ÖNCE çalışır.
- `build_templates.py` — imaj derleme sırasında çalışır:
  `themes/templates/docx/base.docx` (placeholder'lı iskelet) ve
  `themes/templates/pptx/basic.pptx` (10 slayt) üretir.
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
- Agent kuralı: tema dışı renk/font üretmek yasak; tema seçimi final
  raporda belirtilir.

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
│   ├── uploads/                # kullanıcı kaynakları (salt okunur)
│   ├── scratch/                # çalışma alanı
│   └── outputs/                # TESLİMATLAR
├── scripts/
│   ├── validate.py             # yapısal + XSD doğrulama
│   ├── readback.py             # içerik geri okuma
│   ├── recalc_check.py         # xlsx formül yeniden-hesap QA'sı
│   ├── visual_qa.py            # görsel QA eklentisi (varsayılan kapalı)
│   ├── repair_pptx.py          # pptxgenjs XML kusurlarını düzeltir (idempotent)
│   ├── build_templates.py      # imaj derleme: docx/pptx baz şablonları
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
(xlsx: oluştur→validate→readback→recalc, pdf: html→chromium→validate→readback):

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

- Ajan **non-root** (`agent`, uid 1000) kullanıcı; konteyner dışına yazamaz.
- `/opt/docgen/` salt okunur kullanım kuralı; uploads salt okunur.
- Konteyner **tek seferlik** (`--rm`): oturum state'i birikmez, her istek
  temiz sandbox'ta başlar.
- Gotenberg host portu **açılmaz**; yalnızca compose iç ağından erişilir.
- Model anahtarı yalnızca `.env` (gitignore) içinde; konteynere env olarak
  enjekte edilir, hiçbir dosyaya yazılmaz.
- XML işlemlerinde defusedxml (XXE koruması); ofis dosyaları zip-bütünlüğü
  kontrolünden geçer.
