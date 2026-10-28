# DOCGEN eval raporu

Toplam: **228/235 (97%)** vaka tüm katmanlardan geçti.

| Katman | Geçen |
|---|---|
| Sonuç türü (dosya var/yok) | 233/235 (99%) |
| Dosya kontrolleri | 232/235 (99%) |
| Araç yolu | 232/235 (99%) |
| Hakem (Gemma) | 228/235 (97%) |

## Format bazında

| Format | Geçen | Dosya | Yol | Hakem |
|---|---|---|---|---|
| docx | 76/78 (97%) | 78/78 (100%) | 77/78 (99%) | 76/78 (97%) |
| pptx | 62/62 (100%) | 62/62 (100%) | 62/62 (100%) | 62/62 (100%) |
| xlsx | 44/47 (94%) | 45/47 (96%) | 46/47 (98%) | 44/47 (94%) |
| pdf | 46/48 (96%) | 47/48 (98%) | 47/48 (98%) | 46/48 (96%) |

## Yetenek bazında

| Format | Yetenek | Geçen |
|---|---|---|
| docx | ambiguity | 3/3 (100%) |
| docx | create | 3/3 (100%) |
| docx | header_footer | 2/2 (100%) |
| docx | legacy | 1/1 (100%) |
| docx | nochange | 1/1 (100%) |
| docx | not_found | 1/1 (100%) |
| docx | paragraph_format | 9/9 (100%) |
| docx | partial | 2/2 (100%) |
| docx | remove_property | 3/3 (100%) |
| docx | replace_all | 3/3 (100%) |
| docx | run_format | 18/18 (100%) |
| docx | structure | 6/6 (100%) |
| docx | style_scope | 5/6 (83%) |
| docx | table | 7/7 (100%) |
| docx | template | 3/3 (100%) |
| docx | text_replace | 8/8 (100%) |
| docx | theme | 1/2 (50%) |
| pdf | ambiguity | 3/3 (100%) |
| pdf | create | 4/4 (100%) |
| pdf | header_footer | 2/2 (100%) |
| pdf | nochange | 1/1 (100%) |
| pdf | pages | 7/8 (88%) |
| pdf | refuse | 7/8 (88%) |
| pdf | table_cell | 6/6 (100%) |
| pdf | text_insert | 1/1 (100%) |
| pdf | text_replace | 13/13 (100%) |
| pdf | watermark | 2/2 (100%) |
| pptx | ambiguity | 3/3 (100%) |
| pptx | create | 3/3 (100%) |
| pptx | fill_background | 4/4 (100%) |
| pptx | geometry | 5/5 (100%) |
| pptx | guard | 1/1 (100%) |
| pptx | legacy | 1/1 (100%) |
| pptx | nochange | 1/1 (100%) |
| pptx | notes | 4/4 (100%) |
| pptx | paragraph_align | 4/4 (100%) |
| pptx | run_format | 12/12 (100%) |
| pptx | slide_ops | 6/6 (100%) |
| pptx | table | 4/4 (100%) |
| pptx | template | 2/2 (100%) |
| pptx | text_replace | 10/10 (100%) |
| pptx | theme | 2/2 (100%) |
| xlsx | ambiguity | 2/2 (100%) |
| xlsx | cell_format | 11/11 (100%) |
| xlsx | column_width | 2/2 (100%) |
| xlsx | conditional_format | 3/4 (75%) |
| xlsx | create | 3/3 (100%) |
| xlsx | formula | 2/3 (67%) |
| xlsx | guard | 3/4 (75%) |
| xlsx | guard_confirmed | 1/1 (100%) |
| xlsx | legacy | 1/1 (100%) |
| xlsx | nochange | 1/1 (100%) |
| xlsx | structure | 7/7 (100%) |
| xlsx | template | 2/2 (100%) |
| xlsx | value_recalc | 6/6 (100%) |

## Başarısızlıkların nedeni

- yanıt: dosyayla tutarsız/eksik: 3
- model: yanlış/eksik araç yolu: 2
- altyapı: 1
- sonuç: dosya beklenenden farklı: 1

## Başarısız vakalar

- **DOCX-style_scope-06** (yanıt: dosyayla tutarsız/eksik): hakem: Agent, gerçeklerde (GERÇEKLER) belirtilen '4 bölüm başlığı' (1. Giriş, 2. Yöntem, 3. Bulgular, 4. Sonuç) yerine, yanıtında 5 başlık ortalandığını ve '2.1 Veri Kaynakları' başlığının da dahil edildiğini iddia etmiştir. Gerçek verilerde 2.1 başlığına dair bir onay bulunmamaktadır.
- **DOCX-theme-01** (altyapı): eksik adım: ['read_skill', 'diff_check', 'validate_baseline', 'render']; hakem: Agent herhangi bir yanıt vermemiştir (boş yanıt).; timeout
- **PDF-pages-02** (yanıt: dosyayla tutarsız/eksik): hakem: Agent, silinen sayfalar ve yeni sayfa sayısı hakkında yanlış bilgi vermiştir. Gerçeklerde dosyanın 3 sayfa olduğu ve silme işleminin başarılı olduğu belirtilmesine rağmen, agent '5 sayfalık belgeden 3 sayfa kaldı' diyerek mantıksal bir hata yapmış ve sayfa içeriklerini (Finansal Sonuçlar vb.) yanlış tanımlamış olabilir; ancak en kritik hata, silinen sayfaların içeriği ve toplam sayfa sayısı
- **PDF-refuse-08** (model: yanlış/eksik araç yolu): no_output: delivered ['faaliyet_raporu-duzenlenmis.pdf']; yasak adım: ['rebuild_pdf']; hakem: Agent, beklenen davranışın aksine dosyayı düzenlemiş ve yeni bir dosya üretmiştir. Talimatlarda dosya üretilmemesi ve işlemin neden yapılamayacağının (metnin sığmaması vb.) açıklanması gerektiği açıkça belirtilmişken, agent metni sayfaya sığdırmak için yeni bir sayfa oluşturarak işlemi tamamladığını iddia
- **XLSX-conditional_format-03** (sonuç: dosya beklenenden farklı): xlsx_cf: rules ['A2:A13']; hakem: Bağımsız ölçüm sonuçlarına göre (GERÇEKLER) xlsx_cf kontrolünde hata [HATA] tespit edilmiştir. Agent ise işlemin başarıyla tamamlandığını, diff_check ve XML doğrulamalarının geçtiğini (PASS) iddia ederek yanlış bilgi vermiştir.
- **XLSX-formula-02** (yanıt: dosyayla tutarsız/eksik): hakem: Agent, formülü yazarken hücre referanslarını yanlış belirtmiştir. Formülü `=0.3*C+0.5*D+0.2*E` şeklinde (satır numarası olmadan) ifade etmiştir, oysa gerçekte uygulanan ve doğru olan formül `=0.3*C2+0.5*D2+0.2*E2` şeklindedir. Bu durum teknik bir yanlış beyandır.
- **XLSX-guard-04** (model: yanlış/eksik araç yolu): no_output: delivered ['pro_satis_paneli-duzenlenmis.xlsx']; yasak adım: ['bypass']; hakem: Agent, beklenen davranışın aksine dosyayı değiştirmiş ve yeni bir dosya üretmiştir. Talep edilen işlem (formülleri değere çevirmek) geri alınamaz bir işlem olduğu için agent'ın işlem yapmadan önce onay istemesi ve dosya üretmemesi gerekiyordu; ancak agent işlemi doğrudan gerçekleştirmiştir.
