#!/usr/bin/env python3
"""Generate evals/cases.csv — the DOCGEN capability eval set.

Every case is grounded in the fixtures: targets are verified to exist (and,
where the case needs it, to be unique) while the CSV is built, and package
part names used by the checks are read from the fixture itself.

Columns: id, format, capability, fixtures, request, expect, checks, ref,
required, forbidden, judge
  expect   edit | create | refuse | partial | nochange
  checks   JSON list of {"t": <check name in checks.py>, ...args}
  ref      JSON reference solution (reference.py) used by the self-test
  required / forbidden   trajectory step keys (evaluate.py STEPS / FORBID)
"""
import csv
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "lib"))
sys.path.insert(0, str(HERE))
from docgen import docx as D, pdf as F, pptx as P, xlsx as X  # noqa: E402
from docgen.package import Package  # noqa: E402

FX = HERE / "fixtures"
CASES = []

EDIT = {"docx": ["read_skill", "locate", "docgen_docx", "diff_check", "validate_baseline"],
        "pptx": ["read_skill", "locate", "docgen_pptx", "diff_check", "validate_baseline"],
        "xlsx": ["read_skill", "locate", "docgen_xlsx", "diff_check", "validate_baseline"],
        "pdf": ["read_skill", "locate", "pdf_tools", "diff_check"]}
NO_REBUILD = {"docx": ["oxide_edit", "rebuild_md", "hand_zip"],
              "pptx": ["oxide_edit", "pptxgen_edit", "hand_zip"],
              "xlsx": ["oxide_edit", "hand_zip"],
              "pdf": ["crop_route", "rebuild_pdf"]}
PAGE_STEPS = ["read_skill", "pdf_tools", "diff_check"]
THEME_STEPS = ["read_skill", "theme", "diff_check", "validate_baseline"]
J_EDIT = ("Yanıt yapılan değişikliği doğru ve özlü anlatmalı; dosyada olmayan bir değişikliği yapılmış gibi "
          "anlatmamalı; uygulanan değerleri (renk, font, sayı vb.) doğru söylemeli; çıktı dosyasının yolunu vermeli.")
J_CREATE = "Yanıt üretilen dosyanın yolunu ve içeriğini doğru özetlemeli; olmayan bir özelliği iddia etmemeli."
J_NOCHANGE = "Yanıt dosyanın değiştirilmeden teslim edildiğini söylemeli; değişiklik yapıldığını iddia etmemeli."


def add(fmt, cap, fixtures, request, expect, checks, ref=None, required=None, forbidden=None, judge=None):
    n = sum(1 for c in CASES if c["format"] == fmt and c["capability"] == cap) + 1
    if required is None:
        required = {"edit": EDIT[fmt], "partial": EDIT[fmt], "nochange": [],
                    "refuse": ["read_skill"], "create": ["read_skill", "validate"]}[expect]
    if fmt != "pdf" and expect in ("edit", "partial", "create"):
        required = list(required) + ["render"]  # every delivered office file must pass render_check
    if forbidden is None:
        forbidden = NO_REBUILD[fmt] + (["bypass"] if expect in ("refuse", "edit", "partial") else [])
    forbidden = list(forbidden) + ["install"]
    CASES.append({"id": f"{fmt.upper()}-{cap}-{n:02d}", "format": fmt, "capability": cap,
                  "fixtures": ";".join(fixtures), "request": request, "expect": expect,
                  "checks": json.dumps(checks, ensure_ascii=False), "ref": json.dumps(ref or {"fn": "none"}, ensure_ascii=False),
                  "required": ";".join(required), "forbidden": ";".join(forbidden),
                  "judge": judge or {"edit": J_EDIT, "partial": J_EDIT, "create": J_CREATE, "nochange": J_NOCHANGE}.get(expect, "")})


def out(ext):
    return {"t": "output", "ext": ext}


def refuse_judge(why):
    return ("Agent dosyayı değiştirmemeli ve çıktı dosyası üretmemeli. Yanıt işin neden yapılmadığını açıkça söylemeli "
            f"({why}); başarı iddia etmemeli.")


# =============================================================================
# DOCX
# =============================================================================
def docx_unique(f, text):
    return len(D.find(Package(FX / f), text)) == 1


RUN_TARGETS = [
    ("bicimli.docx", "kritik sistemler"), ("bicimli.docx", "Önemli uyarı"), ("bicimli.docx", "dolandırıcılık riski"),
    ("bicimli.docx", "Vurgulu kısım"), ("basliklar.docx", "uzaktan çalışma uygulamasını"),
    ("basliklar.docx", "Anket ve görüşme yöntemleri"), ("basliklar.docx", "Çalışan memnuniyeti arttı"),
    ("satis_raporu.docx", "12.450.000 TL"), ("satis_raporu.docx", "Marmara bölgesi"),
    ("satis_raporu.docx", "hedefin altında kaldı"), ("yapay_zeka_raporu.docx", "yapay zekâ kullanımının"),
    ("yapay_zeka_raporu.docx", "veri güvenliği"), ("yapay_zeka_raporu.docx", "kullanım günlükleri"),
    ("tekrarli_baslik.docx", "üç aşamalı"), ("tekrarli_baslik.docx", "süreç süreleri kısaldı"),
    ("proje_raporu_dagnik.docx", "planlanan takvime"), ("proje_raporu_dagnik.docx", "%78 seviyesinde"),
    ("proje_raporu_dagnik.docx", "Ödeme altyapısı"),
]
RUN_PROPS = [
    ("'{t}' ifadesini koyu yeşil yap.", {"color": "darkgreen"}, {"color": "006400"}),
    ("'{t}' ifadesinin rengini #C00000 yap.", {"color": "C00000"}, {"color": "C00000"}),
    ("'{t}' ifadesini kalın yap.", {"bold": True}, {"bold": True}),
    ("'{t}' ifadesini italik yap.", {"italic": True}, {"italic": True}),
    ("'{t}' ifadesinin altını çiz.", {"underline": True}, {"underline": True}),
    ("'{t}' ifadesini 14 punto yap.", {"size": 14}, {"size": 14}),
    ("'{t}' ifadesinin yazı tipini Arial yap.", {"font": "Arial"}, {"font": "Arial"}),
    ("'{t}' ifadesini sarı ile vurgula.", {"highlight": "yellow"}, {"highlight": "yellow"}),
    ("'{t}' ifadesinin üstünü çiz.", {"strike": True}, {"strike": True}),
    ("'{t}' ifadesini lacivert ve kalın yap.", {"color": "navy", "bold": True}, {"color": "1F3864", "bold": True}),
    ("'{t}' ifadesinin yazı tipini Georgia, boyutunu 12 punto yap.", {"font": "Georgia", "size": 12},
     {"font": "Georgia", "size": 12}),
    ("'{t}' ifadesinin rengini #1F4E79 yap.", {"color": "1F4E79"}, {"color": "1F4E79"}),
]
ALLPROPS = ["bold", "italic", "underline", "strike", "color", "size", "font", "highlight"]
import checks as _C  # noqa: E402


class _Ctx:
    def __init__(self, path):
        self.original = self.output = str(path)


for i, (f, t) in enumerate(RUN_TARGETS):
    assert docx_unique(f, t), (f, t)
    k = i
    while _C.check_docx_chars(_Ctx(FX / f), t, RUN_PROPS[k % len(RUN_PROPS)][1])[0]:
        k += 1  # the original already has this property: pick the next one
    phrase, want, setp = RUN_PROPS[k % len(RUN_PROPS)]
    add("docx", "run_format", [f], f"uploads/{f} dosyasında " + phrase.format(t=t), "edit",
        [out("docx"), {"t": "opens"}, {"t": "docx_chars", "text": t, "props": want, "keep": [p for p in ALLPROPS if p not in want]},
         {"t": "docx_others_same", "allow": [t]}, {"t": "parts_only", "allow": ["word/document.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "run_props", "text": t, "props": setp}]},
        required=EDIT["docx"] + ["docgen_docx"])

PARA = [
    ("bicimli.docx", "Önemli uyarı", "'Önemli uyarı' ile başlayan paragrafı sola yasla.", {"align": "left"}),
    ("bicimli.docx", "Bu paragraf düz metindir", "Ek Notlar bölümündeki paragrafın satır aralığını 1,5 yap.", {"line_spacing": 1.5}),
    ("satis_raporu.docx", "Marmara bölgesi cironun", "Değerlendirme bölümündeki paragrafı iki yana yasla.", {"align": "justify"}),
    ("yapay_zeka_raporu.docx", "Kurumsal bir yönetişim", "4. Sonuç bölümündeki paragrafı ortala.", {"align": "center"}),
    ("tekrarli_baslik.docx", "Program, üç aşamalı", "Giriş bölümündeki paragrafı sağa yasla.", {"align": "right"}),
    ("proje_raporu_dagnik.docx", "Bütçe kullanımı", "'Bütçe kullanımı' ile başlayan paragrafın satır aralığını 2 yap.", {"line_spacing": 2}),
    ("satis_raporu.docx", "Kaynak: Şirket içi", "Kaynakça bölümündeki satırı sağa yasla.", {"align": "right"}),
    ("yapay_zeka_raporu.docx", "Veriler anket ve sistem", "2. Yöntem bölümündeki paragrafın sonrasına 12 punto boşluk bırak.",
     {"space_after": 12}),
]
for f, contains, req, props in PARA:
    refp = {"align": props.get("align"), "line_spacing": props.get("line_spacing"), "space_after": props.get("space_after")}
    add("docx", "paragraph_format", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("docx"), {"t": "docx_para", "contains": contains, **props}, {"t": "docx_others_same", "allow": [contains]},
         {"t": "parts_only", "allow": ["word/document.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "para_props", "contains": contains,
                                                 "props": {k: v for k, v in refp.items() if v is not None}}]})
add("docx", "paragraph_format", ["basliklar.docx"], "uploads/basliklar.docx dosyasında yalnızca 'Giriş' başlığını ortala.", "edit",
    [out("docx"), {"t": "docx_para", "contains": "Giriş", "align": "center"},
     {"t": "docx_para", "contains": "Yöntem", "align": "left"}, {"t": "docx_others_same", "allow": ["Giriş"]},
     {"t": "parts_only", "allow": ["word/document.xml", "word/styles.xml"]}],
    {"fn": "docx_ops", "file": "basliklar.docx", "ops": [{"op": "para_props", "contains": "Giriş", "headings_only": True,
                                                         "props": {"align": "center"}}]})

STYLE = [
    ("basliklar.docx", "Bütün başlıkların rengini #7B2D8E yap.", [0, 1, 2], {"color": "7B2D8E"}, {"color": "7B2D8E"}),
    ("altcizili_basliklar.docx", "Bütün 2. düzey başlıkları lacivert yap; altı çizgileri kalsın.", [2],
     {"color": "navy", "underline": True}, {"color": "1F3864"}),
    ("tekrarli_baslik.docx", "Tüm 1. düzey başlıkların yazı tipini Georgia yap.", [1], {"font": "Georgia"}, {"font": "Georgia"}),
    ("satis_raporu.docx", "Bölüm başlıklarını (1. düzey) 16 punto yap; ana başlığa dokunma.", [1], {"size": 16}, {"size": 16}),
    ("bicimli.docx", "Belgenin ana başlığını (Title) 28 punto yap.", [0], {"size": 28}, {"size": 28}),
]
for f, req, levels, want, setp in STYLE:
    add("docx", "style_scope", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("docx"), {"t": "docx_headings", "levels": levels, "props": want},
         {"t": "parts_only", "allow": ["word/styles.xml", "word/document.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "style_levels", "levels": levels, "props": setp}]})
add("docx", "style_scope", ["yapay_zeka_raporu.docx"], "uploads/yapay_zeka_raporu.docx dosyasında tüm bölüm başlıklarını ortala.",
    "edit", [out("docx")] + [{"t": "docx_para", "contains": h, "align": "center"} for h in
                             ("1. Giriş", "2. Yöntem", "3. Bulgular", "4. Sonuç")] +
    [{"t": "docx_para", "contains": "Bu rapor, kurumda", "align": "left"},
     {"t": "parts_only", "allow": ["word/styles.xml", "word/document.xml"]}],
    {"fn": "docx_ops", "file": "yapay_zeka_raporu.docx", "ops": [{"op": "style_levels", "levels": [1, 2], "props": {"align": "center"}}]})

add("docx", "remove_property", ["bicimli.docx"],
    "uploads/bicimli.docx dosyasında 'Sonuç olarak dolandırıcılık' ile başlayan paragraftaki kalınlığı kaldır, altı çizili kalsın.",
    "edit", [out("docx"), {"t": "docx_chars", "text": "Sonuç olarak dolandırıcılık riski düşüktür", "props": {"bold": False, "underline": True},
                           "keep": ["size", "color"]}, {"t": "docx_others_same", "allow": ["Sonuç olarak dolandırıcılık"]},
             {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "bicimli.docx", "ops": [{"op": "para_runs", "contains": "Sonuç olarak dolandırıcılık", "props": {"bold": False}}]})
add("docx", "remove_property", ["bicimli.docx"],
    "uploads/bicimli.docx dosyasında Özet bölümündeki 'risk analizi' ifadesinde geçen 'risk' kelimesinin altını çizgisini kaldır; rengi ve kalınlığı kalsın.",
    "edit", [out("docx"), {"t": "docx_chars", "text": "risk", "near": "risk analizi", "occurrence": 1,
                           "props": {"underline": False, "bold": True}, "keep": ["color"]},
             {"t": "docx_others_same", "allow": ["risk analizi"]}, {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "bicimli.docx", "ops": [{"op": "run_props", "text": "risk", "near": "risk analizi", "whole_word": True,
                                                       "section": "Özet", "props": {"underline": False}}]})
add("docx", "remove_property", ["altcizili_basliklar.docx"],
    "uploads/altcizili_basliklar.docx dosyasında 2. düzey başlıkların altı çizgisini kaldır.", "edit",
    [out("docx"), {"t": "docx_headings", "levels": [2], "props": {"underline": False}},
     {"t": "parts_only", "allow": ["word/styles.xml", "word/document.xml"]}],
    {"fn": "docx_ops", "file": "altcizili_basliklar.docx", "ops": [{"op": "style_levels", "levels": [2], "props": {"underline": False}}]})

REPL = [
    ("yapay_zeka_raporu.docx", "4. Sonuç", "önerilir", "zorunludur", "Kurumsal bir yönetişim çerçevesi zorunludur."),
    ("satis_raporu.docx", "Özet", "12.450.000 TL", "12.600.000 TL", "toplam ciro 12.600.000 TL"),
    ("satis_raporu.docx", "Kaynakça", "Ocak 2026", "Şubat 2026", "Şubat 2026 çıktısı"),
    ("basliklar.docx", "Yöntem > Sonuç", "katılım yüksek", "katılım çok yüksek", "Yöntemin sonucu: katılım çok yüksek."),
    ("tekrarli_baslik.docx", "3. Sonuç", "kısaldı", "belirgin biçimde kısaldı", "süreç süreleri belirgin biçimde kısaldı"),
    ("proje_raporu_dagnik.docx", None, "%78", "%81", "Bütçe kullanımı %81 seviyesinde"),
    ("yapay_zeka_raporu.docx", "1. Giriş", "mevcut durumunu", "güncel durumunu", "kullanımının güncel durumunu"),
]
for f, sec, old, new, expect_text in REPL:
    where = f"{sec} bölümündeki " if sec else ""
    add("docx", "text_replace", [f], f"uploads/{f} dosyasında {where}'{old}' ifadesini '{new}' olarak değiştir.", "edit",
        [out("docx"), {"t": "text_has", "texts": [expect_text]}, {"t": "text_not", "texts": [old] if old not in new else []},
         {"t": "docx_others_same", "allow": [old, new]}, {"t": "parts_only", "allow": ["word/document.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "replace", "old": old, "new": new,
                                                 "section": sec.split(" > ")[-1] if sec else None}]})
add("docx", "text_replace", ["bicimli.docx"],
    "uploads/bicimli.docx dosyasında Ek Notlar bölümündeki 'risk' kelimesini 'tehlike' yap; belgedeki diğer 'risk'ler aynı kalsın.",
    "edit", [out("docx"), {"t": "text_has", "texts": ["tehlike burada da geçer", "risk analizi", "riski düşüktür", "risk izleniyor"]},
             {"t": "docx_others_same", "allow": ["burada da geçer"]}, {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "bicimli.docx", "ops": [{"op": "replace", "old": "risk", "new": "tehlike", "section": "Ek Notlar", "whole_word": True}]})
for f, old, new, part in (("kurumsal_sablon.docx", "OPPY Kurumsal — Dahili", "OPPY Kurumsal — Gizli", "word/header*.xml"),
                          ("yapay_zeka_raporu.docx", "Yapay Zekâ Raporu — Dahili", "Yapay Zekâ Raporu — Gizli", "word/header*.xml")):
    add("docx", "header_footer", [f], f"uploads/{f} dosyasında sayfa üst bilgisindeki '{old}' yazısını '{new}' yap.", "edit",
        [out("docx"), {"t": "text_has", "texts": [new]}, {"t": "text_not", "texts": [old]},
         {"t": "docx_others_same", "allow": [old, new]}, {"t": "parts_only", "allow": [part]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "replace", "old": old, "new": new, "part": "header"}]})

add("docx", "replace_all", ["bicimli.docx"], "uploads/bicimli.docx dosyasında bütün 'risk' kelimelerini 'tehdit' yap.", "edit",
    [out("docx"), {"t": "text_count", "text": "tehdit", "n": 3}, {"t": "text_not", "texts": ["tehditi"]},
     {"t": "docx_chars", "text": "tehdit", "near": "analizi", "props": {"bold": True, "underline": True, "color": "C00000"}},
     {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "bicimli.docx", "ops": [{"op": "replace_all", "old": "risk", "new": "tehdit"}]},
    judge=J_EDIT + " Kelime içinde geçen 'riski' gibi çekimli biçimlere dokunulmadıysa bunu belirtmesi beklenir.")
add("docx", "replace_all", ["satis_raporu.docx"], "uploads/satis_raporu.docx dosyasında '2025' geçen her yeri '2026' yap.", "edit",
    [out("docx"), {"t": "text_not", "texts": ["2025"]}, {"t": "text_has", "texts": ["2026 Satış Performans Raporu", "2026 yılında"]},
     {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "satis_raporu.docx", "ops": [{"op": "replace_all", "old": "2025", "new": "2026"}]})
add("docx", "replace_all", ["basliklar.docx"], "uploads/basliklar.docx dosyasında her iki 'Sonuç' başlığının adını 'Değerlendirme' yap.",
    "edit", [out("docx"), {"t": "docx_heading_texts", "absent": ["Sonuç"]}, {"t": "text_count", "text": "Değerlendirme", "n": 2},
             {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "basliklar.docx", "ops": [{"op": "replace_all", "old": "Sonuç", "new": "Değerlendirme", "headings_only": True}]})

TBL = [
    ("Bölge Sonuçları tablosunda Ege'nin cirosunu 3.250.000 yap.", "satis_raporu.docx",
     [{"t": "docx_table_cell", "table_has": "Ege", "row": "Ege", "col": "Ciro", "value": "3.250.000"},
      {"t": "docx_table_cell", "table_has": "Ege", "row": "Ege", "col": "Pay", "value": "24,9"},
      {"t": "text_has", "texts": ["12.450.000 TL"]}],
     [{"op": "table_cell", "table_has": "Ege", "row": "Ege", "col": "Ciro", "value": "3.250.000"}]),
    ("Bölge Sonuçları tablosunda Karadeniz'in payını 15,0 yap.", "satis_raporu.docx",
     [{"t": "docx_table_cell", "table_has": "Ege", "row": "Karadeniz", "col": "Pay", "value": "15,0"}],
     [{"op": "table_cell", "table_has": "Ege", "row": "Karadeniz", "col": "Pay", "value": "15,0"}]),
    ("Bölge Sonuçları tablosuna Karadeniz'in altına 'Akdeniz | 900.000 | 7,2' satırını ekle.", "satis_raporu.docx",
     [{"t": "docx_table_shape", "table_has": "Ege", "rows": 6, "row_values": ["Akdeniz", "900.000", "7,2"]}],
     [{"op": "table_add_row", "table_has": "Ege", "after": "Karadeniz", "values": ["Akdeniz", "900.000", "7,2"]}]),
    ("Bölge Sonuçları tablosunun sonuna 'Hedef' başlıklı boş bir sütun ekle.", "satis_raporu.docx",
     [{"t": "docx_table_shape", "table_has": "Ege", "cols": 4, "header": "Hedef"}],
     [{"op": "table_add_col", "table_has": "Ege", "header": "Hedef"}]),
    ("Veri kaynakları tablosuna 'Pay (%)' başlıklı boş bir sütun ekle.", "yapay_zeka_raporu.docx",
     [{"t": "docx_table_shape", "table_has": "Ticket", "cols": 3, "header": "Pay (%)"}, {"t": "docx_media", "count": 3}],
     [{"op": "table_add_col", "table_has": "Ticket", "header": "Pay (%)"}]),
    ("Bölge Sonuçları tablosundan 'Pay (%)' sütununu sil.", "satis_raporu.docx",
     [{"t": "docx_table_shape", "table_has": "Ege", "cols": 2}, {"t": "text_has", "texts": ["3.100.000"]}],
     [{"op": "table_del_col", "table_has": "Ege", "col": "Pay"}]),
    ("tablonun başlık satırını lacivert dolguyla renklendir.", "bicimli.docx",
     [{"t": "docx_cell_fill", "table_has": "Ölçüt", "row": 0, "color": "navy"}],
     [{"op": "table_header_fill", "table_has": "Ölçüt", "color": "1F3864"}]),
]
for req, f, checks, ops in TBL:
    add("docx", "table", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("docx")] + checks + [{"t": "parts_only", "allow": ["word/document.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": ops})

add("docx", "structure", ["yapay_zeka_raporu.docx"],
    "uploads/yapay_zeka_raporu.docx dosyasında 2.1 Veri Kaynakları alt başlığını, altındaki paragraf ve tabloyla birlikte sil.", "edit",
    [out("docx"), {"t": "docx_heading_texts", "absent": ["2.1 Veri Kaynakları"], "present": ["2. Yöntem", "3. Bulgular"]},
     {"t": "text_not", "texts": ["Üç kaynak kullanıldı", "Ticket kayıtları"]}, {"t": "text_has", "texts": ["Veriler anket ve sistem"]},
     {"t": "docx_media", "count": 3},
     {"t": "docx_others_same", "allow_delete": ["2.1 Veri", "Üç kaynak", "Kaynak", "Kayıt sayısı", "Anket", "1.240", "Ticket", "5.870"]},
     {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "yapay_zeka_raporu.docx", "ops": [{"op": "delete_section", "heading": "2.1 Veri Kaynakları"}]})
add("docx", "structure", ["tekrarli_baslik.docx"], "uploads/tekrarli_baslik.docx dosyasında genel uygulamanın sonuç bölümünü sil.",
    "edit", [out("docx"), {"t": "docx_heading_texts", "absent": ["5. Sonuç"], "present": ["3. Sonuç"]},
             {"t": "text_not", "texts": ["nihai sonucu"]}, {"t": "text_has", "texts": ["süreç süreleri kısaldı"]},
             {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "tekrarli_baslik.docx", "ops": [{"op": "delete_section", "heading": "5. Sonuç"}]})
add("docx", "structure", ["satis_raporu.docx"],
    "uploads/satis_raporu.docx dosyasında Değerlendirme bölümündeki paragrafı şu metinle değiştir: 'Marmara ve Ege bölgeleri hedefi aştı; Karadeniz için yeni bir saha planı hazırlanıyor.'",
    "edit", [out("docx"), {"t": "text_has", "texts": ["Karadeniz için yeni bir saha planı hazırlanıyor"]},
             {"t": "text_not", "texts": ["Marmara bölgesi cironun"]}, {"t": "docx_heading_texts", "present": ["Değerlendirme"]},
             {"t": "docx_others_same", "allow": ["Marmara bölgesi cironun", "saha planı"]},
             {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "satis_raporu.docx", "ops": [{"op": "set_para", "heading": "Değerlendirme",
                                                            "text": "Marmara ve Ege bölgeleri hedefi aştı; Karadeniz için yeni bir saha planı hazırlanıyor."}]})
for f, heading, nxt, text in (
        ("yapay_zeka_raporu.docx", "3. Bulgular", "4. Sonuç", "Eğitim ihtiyacı tüm birimlerde yüksek bulundu."),
        ("tekrarli_baslik.docx", "2. Pilot Uygulama", "3. Sonuç", "Pilot süresince haftalık ilerleme raporları hazırlandı."),
        ("pro_kurumsal_rapor.docx", "3. Operasyon", "4. Sonuç ve Öneriler", "Depo otomasyonu 2026'nın ikinci çeyreğinde devreye alınacak.")):
    add("docx", "structure", [f], f"uploads/{f} dosyasında {heading} bölümünün sonuna şu paragrafı ekle: '{text}'", "edit",
        [out("docx"), {"t": "docx_before", "text": text, "next_heading": nxt},
         {"t": "docx_others_same", "allow_insert": [text]},
         {"t": "parts_only", "allow": ["word/document.xml", "word/numbering.xml"]}],
        {"fn": "docx_ops", "file": f, "ops": [{"op": "append_section", "heading": heading, "text": text}]})

add("docx", "theme", ["yapay_zeka_raporu.docx"], "uploads/yapay_zeka_raporu.docx dosyasında belgenin tema renklerini yeşil tonlarına çevir; kendi rengi olan yazılar korunsun.",
    "edit", [out("docx"), {"t": "docx_theme", "slot": "accent1", "color": "green"},
             {"t": "docx_chars", "text": "kritik", "props": {"color": "C00000"}}, {"t": "docx_media", "count": 3},
             {"t": "parts_only", "allow": ["word/theme/theme1.xml", "word/styles.xml", "word/stylesWithEffects.xml", "word/document.xml"]}],
    {"fn": "docx_ops", "file": "yapay_zeka_raporu.docx", "ops": [{"op": "theme_color", "slot": s, "color": c} for s, c in
                                                                 (("accent1", "2E7D32"), ("accent2", "66BB6A"), ("dk2", "1B5E20"))]},
    required=THEME_STEPS)
add("docx", "theme", ["basliklar.docx"], "uploads/basliklar.docx dosyasında tema yazı tiplerini başlıklar için Georgia, gövde için Calibri yap.",
    "edit", [out("docx"), {"t": "docx_theme", "major": "Georgia", "minor": "Calibri"},
             {"t": "parts_only", "allow": ["word/theme/theme1.xml", "word/styles.xml", "word/stylesWithEffects.xml"]}],
    {"fn": "docx_ops", "file": "basliklar.docx", "ops": [{"op": "theme_fonts", "major": "Georgia", "minor": "Calibri"}]},
    required=THEME_STEPS)

for f, req, cands in (("basliklar.docx", "'Sonuç' başlığını kırmızı yap.", "Yöntem altındaki ve Bulgular altındaki iki Sonuç başlığı"),
                      ("tekrarli_baslik.docx", "Sonuç bölümünü sil.", "3. Sonuç ve 5. Sonuç"),
                      ("basliklar.docx", "Sonuç bölümünün sonuna 'Not: Veriler 2026'da güncellenecek.' cümlesini ekle.",
                       "Yöntem altındaki ve Bulgular altındaki iki Sonuç bölümü")):
    add("docx", "ambiguity", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge("hedef belirsiz") + f" Yanıt adayları ({cands}) ayırt edilebilir biçimde listelemeli ve hangisinin kastedildiğini sormalı.")
add("docx", "not_found", ["bicimli.docx"], "uploads/bicimli.docx dosyasında 'Ekler' bölümünü sil.", "refuse", [{"t": "no_output"}],
    judge=refuse_judge("belgede 'Ekler' adlı bir bölüm yok"))
add("docx", "partial", ["satis_raporu.docx"],
    "uploads/satis_raporu.docx dosyasında Özet bölümünün adını 'Yönetici Özeti' yap ve 'Lojistik Notları' bölümünü sil.", "partial",
    [{"t": "docx_heading_texts", "present": ["Yönetici Özeti"]}, {"t": "docx_others_same", "allow": ["Özet"]}],
    {"fn": "docx_ops", "file": "satis_raporu.docx", "ops": [{"op": "rename_heading", "old": "Özet", "new": "Yönetici Özeti"}]},
    judge=J_EDIT + " 'Lojistik Notları' bölümü belgede olmadığı için silinmediğini açıkça söylemeli. (Hiç dosya üretmeyip bunu "
                   "gerekçesiyle söylemesi de kabul edilir.)")
add("docx", "partial", ["satis_raporu.docx"],
    "uploads/satis_raporu.docx dosyasında '2025' yazan yerleri '2026' yap, '2019' yazan yerleri de '2020' yap.", "partial",
    [{"t": "text_not", "texts": ["2025"]}, {"t": "parts_only", "allow": ["word/document.xml"]}],
    {"fn": "docx_ops", "file": "satis_raporu.docx", "ops": [{"op": "replace_all", "old": "2025", "new": "2026"}]},
    judge=J_EDIT + " Belgede '2019' geçmediği için o kısmın yapılmadığını açıkça söylemeli.")

for f, req, title, secs, checks in (
        ("pro_kurumsal_sablon.dotx", "Bu Word şablonunu kullanarak 'Uzaktan Çalışma Politikası' başlıklı; Amaç, Kapsam ve Kurallar bölümlerinden oluşan kısa bir belge hazırla.",
         "Uzaktan Çalışma Politikası", ["Amaç", "Kapsam", "Kurallar"],
         [{"t": "docx_media", "header_image": True}, {"t": "text_has", "texts": ["Kuzey Lojistik A.Ş. · Kurumsal Belge"]},
          {"t": "text_not", "texts": ["Belge başlığı"]}]),
        ("kurumsal_sablon.docx", "Bu şablonla 'Toplantı Tutanağı' başlıklı; Katılımcılar, Gündem ve Kararlar bölümleri olan bir belge hazırla. Şablondaki örnek metni kaldır.",
         "Toplantı Tutanağı", ["Katılımcılar", "Gündem", "Kararlar"],
         [{"t": "text_has", "texts": ["OPPY Kurumsal — Dahili", "Gizli"]}, {"t": "text_not", "texts": ["Bu dosya bir şablon örneğidir."]}]),
        ("pro_kurumsal_sablon.dotx", "Bu şablonla 'Ziyaretçi Kabul Prosedürü' başlıklı; Giriş, Adımlar ve Sorumluluklar bölümleri olan bir belge oluştur.",
         "Ziyaretçi Kabul Prosedürü", ["Giriş", "Adımlar", "Sorumluluklar"],
         [{"t": "docx_media", "header_image": True}])):
    add("docx", "template", [f], f"uploads/{f} — {req}", "create",
        [out("docx"), {"t": "opens"}, {"t": "text_has", "texts": [title] + secs},
         {"t": "docx_heading_texts", "present": secs}, {"t": "docx_style_used", "min_heading_paras": 3}] + checks,
        {"fn": "docx_template", "file": f, "title": title, "sections": secs},
        required=["read_skill", "template", "validate_baseline"], forbidden=["bypass"])
add("docx", "create", [], "Q3 2026 satış özeti için bir Word belgesi oluştur: başlık; Özet, Bölgeler ve Sonraki Adımlar bölümleri; bölge cirolarını gösteren bir tablo. Bölüm başlıkları koyu kırmızı (#8B0000) olsun.",
    "create", [out("docx"), {"t": "opens"}, {"t": "docx_heading_texts", "present": ["Özet", "Bölgeler", "Sonraki Adımlar"]},
               {"t": "docx_tables", "min_tables": 1}, {"t": "docx_any_color", "color": "8B0000"}],
    {"fn": "docx_create", "name": "q3-satis-ozeti.docx", "title": "Q3 2026 Satış Özeti",
     "sections": ["Özet", "Bölgeler", "Sonraki Adımlar"], "heading_color": "8B0000", "table": True},
    required=["read_skill", "create_docx", "validate"], forbidden=[])
add("docx", "create", [], "İzin politikası için en az üç numaralı maddeden oluşan kısa bir Word belgesi hazırla.", "create",
    [out("docx"), {"t": "opens"}, {"t": "docx_numbered", "min_items": 3}],
    {"fn": "docx_create", "name": "izin-politikasi.docx", "title": "İzin Politikası", "sections": ["Kurallar"], "numbered": True},
    required=["read_skill", "create_docx", "validate"], forbidden=[])
add("docx", "create", [], "minimal temasıyla 'Bilgi Güvenliği Duyurusu' başlıklı, tek sayfalık bir Word duyurusu hazırla.", "create",
    [out("docx"), {"t": "opens"}, {"t": "text_has", "texts": ["Bilgi Güvenliği Duyurusu"]}],
    {"fn": "docx_create", "name": "duyuru.docx", "title": "Bilgi Güvenliği Duyurusu", "sections": ["Duyuru"]},
    required=["read_skill", "create_docx", "validate"], forbidden=[])
add("docx", "legacy", ["bicimli.doc"], "uploads/bicimli.doc dosyasında 'risk analizi' ifadesini kalın yap.", "edit",
    [out("docx"), {"t": "docx_chars", "text": "risk analizi", "props": {"bold": True}}],
    {"fn": "legacy", "file": "bicimli.doc", "target": "docx",
     "then": {"fn": "docx_ops", "ops": [{"op": "run_props", "text": "risk analizi", "props": {"bold": True}}]}},
    required=["read_skill", "legacy", "docgen_docx"], forbidden=["oxide_edit", "bypass"],
    judge=J_EDIT + " Dosyanın .docx biçimine dönüştürüldüğünü belirtmeli.")
add("docx", "nochange", ["yapay_zeka_raporu.docx"], "uploads/yapay_zeka_raporu.docx dosyasında hiçbir şeyi değiştirme, olduğu gibi geri ver.",
    "nochange", [out("docx"), {"t": "identical"}], {"fn": "copy", "file": "yapay_zeka_raporu.docx"})

# =============================================================================
# PPTX
# =============================================================================
def slide_part(f, idx):
    return P.slides(Package(FX / f))[idx - 1].part


def notes_part(f, idx):
    return P.slides(Package(FX / f))[idx - 1].notes_part


PREPL = [
    ("kurumsal_deste.pptx", 6, "Sonuç slaytındaki 'Büyüme sürüyor' maddesini 'Büyüme 2026'da da sürecek'", "Büyüme sürüyor", "Büyüme 2026'da da sürecek", None),
    ("yonetim_sunumu.pptx", 2, "2. slaytın başlığındaki '%11' ifadesini '%12'; maddelere dokunma", "%11", "%12", True),
    ("yonetim_sunumu.pptx", 4, "4. slayttaki 'Ciro 1.800.000 TL' maddesini 'Ciro 1.850.000 TL'", "Ciro 1.800.000 TL", "Ciro 1.850.000 TL", None),
    ("yonetim_sunumu.pptx", 3, "3. slayttaki '1.240 yeni müşteri kazanıldı' maddesini '1.300 yeni müşteri kazanıldı'", "1.240 yeni", "1.300 yeni", None),
    ("egitim_plani.pptx", 3, "Program takvimi slaytındaki 'Hafta 3: Saha çalışması' maddesini 'Hafta 3: Saha ve mentor çalışması'", "Saha çalışması", "Saha ve mentor çalışması", None),
    ("egitim_plani.pptx", 1, "kapak slaytındaki 'Ocak 2026' alt başlığını 'Şubat 2026'", "Ocak 2026", "Şubat 2026", None),
    ("kurumsal_sunum.pptx", 4, "Bütçe dağılımı slaytındaki 'Ar-Ge %25' maddesini 'Ar-Ge %30'", "Ar-Ge %25", "Ar-Ge %30", None),
    ("kurumsal_deste.pptx", 2, "2. slayttaki 'Kuzey bölgesi geride kaldı' maddesini 'Kuzey bölgesi toparlanıyor'", "Kuzey bölgesi geride kaldı", "Kuzey bölgesi toparlanıyor", None),
    ("bicimli.pptx", 2, "2. slayttaki 'risk burada da izlenecek' metnini 'risk haftalık izlenecek'", "risk burada da izlenecek", "risk haftalık izlenecek", None),
    ("yonetim_sunumu.pptx", 6, "Teşekkürler slaytındaki 'Sorular için ekibimize ulaşın' metnini 'Sorularınız için: strateji@ornek.com'", "Sorular için ekibimize ulaşın", "Sorularınız için: strateji@ornek.com", None),
]
for f, idx, req, old, new, in_title in PREPL:
    lacks = [old] if in_title is None else []
    extra = [{"t": "pptx_text", "slide": idx, "has": ["geçen yıla göre %11"]}] if in_title else []
    add("pptx", "text_replace", [f], f"uploads/{f} dosyasında {req} olarak değiştir.", "edit",
        [out("pptx"), {"t": "pptx_text", "slide": idx, "has": [new], "lacks": lacks}] + extra +
        [{"t": "pptx_others_same", "changed_slides": [idx]}, {"t": "parts_only", "allow": [slide_part(f, idx)]}],
        {"fn": "pptx_ops", "file": f, "ops": [{"op": "replace", "slide": idx, "old": old, "new": new, "in_title": in_title}]})

PFMT = [
    ("kurumsal_deste.pptx", 2, "2. slayttaki 'Kaynak: İç analiz' yazısını 10 punto ve italik yap.", "Kaynak: İç analiz", {"size": 10, "italic": True}, None),
    ("bicimli.pptx", 1, "1. slayttaki ' seviyede risk var' kısmını kalın yap.", " seviyede risk var", {"bold": True}, None),
    ("yonetim_sunumu.pptx", 1, "kapak başlığının yazı tipini Georgia yap.", "2025 Yönetim Sunumu", {"font": "Georgia"}, True),
    ("bicimli.pptx", 2, "2. slayttaki 'risk' kelimesini kırmızı yap.", "risk", {"color": "red"}, None),
    ("egitim_plani.pptx", 2, "Eğitim hedefleri slaytındaki 'Ürün bilgisi' maddesini kalın yap.", "Ürün bilgisi", {"bold": True}, None),
    ("kurumsal_sunum.pptx", 3, "Öncelikler slaytındaki 'Dijital kanal' maddesinin rengini #C00000 yap.", "Dijital kanal", {"color": "C00000"}, None),
    ("yonetim_sunumu.pptx", 5, "5. slayttaki 'Dijital kanalı büyüt' maddesini italik yap.", "Dijital kanalı büyüt", {"italic": True}, None),
    ("kurumsal_deste.pptx", 6, "Sonuç slaytındaki 'Büyüme sürüyor' maddesini 28 punto yap.", "Büyüme sürüyor", {"size": 28}, None),
    ("egitim_plani.pptx", 4, "4. slaytın başlığını 40 punto yap.", "Ölçme ve değerlendirme", {"size": 40}, True),
    ("yonetim_sunumu.pptx", 3, "3. slayttaki 'Memnuniyet %87' maddesini yeşil yap.", "Memnuniyet %87", {"color": "green"}, None),
    ("kurumsal_sunum.pptx", 2, "Pazar görünümü slaytının başlığının yazı tipini Arial yap.", "Pazar görünümü", {"font": "Arial"}, True),
    ("bicimli.pptx", 1, "1. slayttaki metin kutusunda geçen 'risk' kelimesinin altını çiz; başlığa dokunma.", "risk", {"underline": True}, False),
]
SETP = {"red": "FF0000", "green": "00B050"}
for f, idx, req, text, want, in_title in PFMT:
    setp = {k: SETP.get(v, v) for k, v in want.items()}
    ops = [{"op": "run_props", "slide": idx, "text": text, "props": setp, "in_title": in_title,
            "whole_word": text == "risk"}]
    add("pptx", "run_format", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("pptx"), {"t": "pptx_chars", "slide": idx, "text": text, "props": want, "in_title": in_title},
         {"t": "pptx_others_same", "changed_slides": [idx]}, {"t": "parts_only", "allow": [slide_part(f, idx)]}],
        {"fn": "pptx_ops", "file": f, "ops": ops})

for f, idx, req, shape_text, align in (
        ("kurumsal_deste.pptx", 2, "2. slayttaki 'Kaynak: İç analiz' yazısını sağa yasla.", "Kaynak: İç analiz", "right"),
        ("yonetim_sunumu.pptx", 6, "Teşekkürler slaytındaki metni ortala.", "Sorular için", "center"),
        ("kurumsal_sunum.pptx", 1, "kapak slaytındaki alt başlığı sağa yasla.", "Yönetim Kurulu", "right"),
        ("egitim_plani.pptx", 2, "Eğitim hedefleri slaytındaki maddeleri ortala.", "Yeni çalışanların", "center")):
    add("pptx", "paragraph_align", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("pptx"), {"t": "pptx_others_same", "changed_slides": [idx]}, {"t": "parts_only", "allow": [slide_part(f, idx)]},
         {"t": "pptx_align", "slide": idx, "shape_text": shape_text, "align": {"right": "r", "center": "ctr"}[align]}],
        {"fn": "pptx_ops", "file": f, "ops": [{"op": "align", "slide": idx, "shape_text": shape_text, "align": align}]})

GEO = [
    ("2. slayttaki 'Önemli' kutusunu sağ üst köşeye taşı.", {"text": "Önemli"}, {"pos": "top-right"}, {"op": "move", "slide": 2, "shape_text": "Önemli", "pos": "top-right"}),
    ("2. slayttaki 'Önemli' kutusunu sol alt köşeye taşı.", {"text": "Önemli"}, {"pos": "bottom-left"}, {"op": "move", "slide": 2, "shape_text": "Önemli", "pos": "bottom-left"}),
    ("2. slayttaki 'Kaynak: İç analiz' metin kutusunu iki kat genişlet; yüksekliği aynı kalsın.", {"text": "Kaynak: İç analiz"}, {"width_factor": 2}, {"op": "widen", "slide": 2, "shape_text": "Kaynak: İç analiz", "factor": 2}),
    ("2. slayttaki resmi yarı boyuta küçült.", {"kind": "pic"}, {"scale": 0.5}, {"op": "scale", "slide": 2, "kind": "pic", "factor": 0.5}),
    ("2. slayttaki resmi oranını koruyarak 1,5 katına büyüt.", {"kind": "pic"}, {"scale": 1.5}, {"op": "scale", "slide": 2, "kind": "pic", "factor": 1.5}),
]
for req, sel, want, op in GEO:
    add("pptx", "geometry", ["kurumsal_deste.pptx"], f"uploads/kurumsal_deste.pptx dosyasında {req}", "edit",
        [out("pptx"), {"t": "pptx_geom", "slide": 2, **sel, **want}, {"t": "pptx_others_same", "changed_slides": [2]},
         {"t": "parts_only", "allow": [slide_part("kurumsal_deste.pptx", 2)]}],
        {"fn": "pptx_ops", "file": "kurumsal_deste.pptx", "ops": [op]})

add("pptx", "fill_background", ["kurumsal_deste.pptx"], "uploads/kurumsal_deste.pptx dosyasında 2. slayttaki 'Önemli' kutusunun dolgusunu #FF8800 yap.",
    "edit", [out("pptx"), {"t": "pptx_fill", "slide": 2, "text": "Önemli", "color": "FF8800"},
             {"t": "pptx_others_same", "changed_slides": [2]}, {"t": "parts_only", "allow": [slide_part("kurumsal_deste.pptx", 2)]}],
    {"fn": "pptx_ops", "file": "kurumsal_deste.pptx", "ops": [{"op": "fill", "slide": 2, "shape_text": "Önemli", "color": "FF8800"}]})
for f, title, color, fam in (("kurumsal_deste.pptx", "Sonuç", "EDEDED", "lightgray"), ("egitim_plani.pptx", "Eğitim hedefleri", "DDEBF7", "lightblue"),
                             ("yonetim_sunumu.pptx", "Teşekkürler", "1F3864", "navy")):
    idx = next(s.index for s in P.slides(Package(FX / f)) if s.title == title)
    word = {"lightgray": "açık gri", "lightblue": "açık mavi", "navy": "lacivert"}[fam]
    add("pptx", "fill_background", [f], f"uploads/{f} dosyasında yalnızca '{title}' slaytının arka planını {word} yap.", "edit",
        [out("pptx"), {"t": "pptx_bg", "title": title, "color": fam}, {"t": "pptx_others_same", "changed_slides": [idx]},
         {"t": "parts_only", "allow": [slide_part(f, idx)]}],
        {"fn": "pptx_ops", "file": f, "ops": [{"op": "bg", "title": title, "color": color}]})

for f, idx, text in (("yonetim_sunumu.pptx", 3, "Yeni müşteriler çoğunlukla dijital kanaldan geldi."),
                     ("egitim_plani.pptx", 4, "Ölçüm quiz ve mentor geri bildirimiyle yapılır."),
                     ("kurumsal_deste.pptx", 2, "Pazar büyümesini rakamlarla anlat."),
                     ("yonetim_sunumu.pptx", 6, "Kapanış ve teşekkür.")):
    add("pptx", "notes", [f], f"uploads/{f} dosyasında {idx}. slaytın konuşmacı notunu '{text}' olarak değiştir.", "edit",
        [out("pptx"), {"t": "pptx_notes", "slide": idx, "equals": text},
         {"t": "pptx_others_same"}, {"t": "parts_only", "allow": [notes_part(f, idx)]}],
        {"fn": "pptx_ops", "file": f, "ops": [{"op": "notes", "slide": idx, "text": text}]})

S5 = slide_part("kurumsal_deste.pptx", 5)
for req, chk, op in (
        ("Ek tablodaki Ege'nin 2025 değerini 5,3 yap.", {"row": "Ege", "col": "2025", "value": "5,3"}, {"op": "table_cell", "slide": 5, "row": "Ege", "col": "2025", "value": "5,3"}),
        ("Ek tablodaki başlık satırını kalın yap ve ortala.", {"row_bold": 0, "row_align": [0, "ctr"]}, {"op": "table_header", "slide": 5, "bold": True, "align": "center"}),
        ("Ek tabloya '2026 Hedef' başlıklı boş bir sütun ekle.", {"cols": 4, "header": "2026 Hedef"}, {"op": "table_add_col", "slide": 5, "header": "2026 Hedef"}),
        ("Ek tablodaki Marmara'nın 2024 hücresini sarı dolguyla işaretle.", {"cell_fill": ["Marmara", "2024", "yellow"]}, {"op": "table_cell_fill", "slide": 5, "row": "Marmara", "col": "2024", "color": "FFFF00"})):
    add("pptx", "table", ["kurumsal_deste.pptx"], f"uploads/kurumsal_deste.pptx dosyasında {req}", "edit",
        [out("pptx"), {"t": "pptx_table", "slide": 5, **chk}, {"t": "pptx_slides", "hidden": ["Ek: Ayrıntılı Tablo"]},
         {"t": "pptx_others_same", "changed_slides": [5]}, {"t": "parts_only", "allow": [S5]}],
        {"fn": "pptx_ops", "file": "kurumsal_deste.pptx", "ops": [op]})

add("pptx", "slide_ops", ["yonetim_sunumu.pptx"], "uploads/yonetim_sunumu.pptx dosyasında Teşekkürler slaytını sil.", "edit",
    [out("pptx"), {"t": "pptx_slides", "count": 5}, {"t": "text_not", "texts": ["Teşekkürler"]},
     {"t": "pptx_others_same", "changed_titles": ["Teşekkürler"]}],
    {"fn": "pptx_ops", "file": "yonetim_sunumu.pptx", "ops": [{"op": "delete_slide", "title": "Teşekkürler"}]})
add("pptx", "slide_ops", ["kurumsal_deste.pptx"], "uploads/kurumsal_deste.pptx dosyasında Süreç slaytını Pazar Görünümü slaytından önceye taşı.", "edit",
    [out("pptx"), {"t": "pptx_slides", "titles": ["Kurumsal Strateji 2026", "Süreç", "Pazar Görünümü", "Gelir Grafiği"]},
     {"t": "pptx_link", "from_title": "Sonuç", "to_title": "Pazar Görünümü"}, {"t": "pptx_others_same"},
     {"t": "parts_only", "allow": ["ppt/presentation.xml"]}],
    {"fn": "pptx_ops", "file": "kurumsal_deste.pptx", "ops": [{"op": "move_slide", "title": "Süreç", "before": "Pazar Görünümü"}]})
add("pptx", "slide_ops", ["egitim_plani.pptx"], "uploads/egitim_plani.pptx dosyasında Program takvimi slaytını çoğalt (kopyası hemen arkasında olsun).", "edit",
    [out("pptx"), {"t": "pptx_slides", "count": 5, "titles": ["Çalışan Eğitim Planı", "Eğitim hedefleri", "Program takvimi", "Program takvimi"]},
     {"t": "pptx_others_same"}],
    {"fn": "pptx_ops", "file": "egitim_plani.pptx", "ops": [{"op": "duplicate", "title": "Program takvimi"}]})
add("pptx", "slide_ops", ["yonetim_sunumu.pptx"], "uploads/yonetim_sunumu.pptx dosyasında Karadeniz slaytını gizle (silme).", "edit",
    [out("pptx"), {"t": "pptx_slides", "count": 6, "hidden": ["Karadeniz"]}, {"t": "pptx_others_same", "changed_titles": ["Karadeniz"]}],
    {"fn": "pptx_ops", "file": "yonetim_sunumu.pptx", "ops": [{"op": "hide", "title": "Karadeniz", "hidden": True}]})
add("pptx", "slide_ops", ["kurumsal_deste.pptx"], "uploads/kurumsal_deste.pptx dosyasında gizli ek tablo slaytını görünür yap.", "edit",
    [out("pptx"), {"t": "pptx_slides", "count": 6, "visible": ["Ek: Ayrıntılı Tablo"]}, {"t": "pptx_others_same", "changed_slides": [5]}],
    {"fn": "pptx_ops", "file": "kurumsal_deste.pptx", "ops": [{"op": "hide", "title": "Ek: Ayrıntılı Tablo", "hidden": False}]})
add("pptx", "slide_ops", ["egitim_plani.pptx"], "uploads/egitim_plani.pptx dosyasında Ölçme ve değerlendirme slaytını sil.", "edit",
    [out("pptx"), {"t": "pptx_slides", "count": 3}, {"t": "pptx_others_same", "changed_titles": ["Ölçme"]}],
    {"fn": "pptx_ops", "file": "egitim_plani.pptx", "ops": [{"op": "delete_slide", "title": "Ölçme ve değerlendirme"}]})
add("pptx", "guard", ["kurumsal_deste.pptx"], "uploads/kurumsal_deste.pptx dosyasında Pazar Görünümü slaytını sil.", "refuse", [{"t": "no_output"}],
    judge=refuse_judge("Sonuç slaytı bu slayta köprü veriyor; silmek köprüyü bozar") + " Onay istenebileceğini belirtmesi iyi olur.")

add("pptx", "theme", ["kurumsal_sunum.pptx"], "uploads/kurumsal_sunum.pptx dosyasında sunumun ana rengini (Accent 1) bordo #800020 yap.", "edit",
    [out("pptx"), {"t": "pptx_theme", "slot": "accent1", "color": "800020"}, {"t": "pptx_others_same"},
     {"t": "parts_only", "allow": ["ppt/theme/theme*.xml"]}],
    {"fn": "pptx_ops", "file": "kurumsal_sunum.pptx", "ops": [{"op": "theme_color", "slot": "accent1", "color": "800020"}]},
    required=THEME_STEPS)
add("pptx", "theme", ["egitim_plani.pptx"], "uploads/egitim_plani.pptx dosyasında sunumun tema başlık yazı tipini Georgia yap.", "edit",
    [out("pptx"), {"t": "pptx_theme", "major": "Georgia"}, {"t": "parts_only", "allow": ["ppt/theme/theme*.xml"]}],
    {"fn": "pptx_ops", "file": "egitim_plani.pptx", "ops": [{"op": "theme_fonts", "major": "Georgia"}]},
    required=THEME_STEPS)

for f, req, cands in (("bicimli.pptx", "'risk' kelimesini kırmızı yap.", "1. slayt ve 2. slayttaki 'risk'"),
                      ("yonetim_sunumu.pptx", "'%11' ifadesini '%12' yap.", "2. slaytın başlığı ve maddesi"),
                      ("kurumsal_sunum.pptx", "Bütçe slaytının başlığını 'Bütçe Planı' yap.", "'Bütçe dağılımı' ve 'Bütçe grafiği' slaytları")):
    add("pptx", "ambiguity", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge("hedef belirsiz") + f" Adayları ({cands}) listeleyip hangisinin kastedildiğini sormalı. "
              "(Kullanıcı açıkça 'her yerde' demediği sürece hepsini değiştirmek doğru değildir.)")

for f, req, titles, forbid in (
        ("kurumsal_sunum.pptx", "Bu sunumu şablon olarak kullan; tasarımını koruyarak 'Yeni Ürün Lansmanı' konulu 4 slaytlık yeni bir sunum hazırla.",
         ["Yeni Ürün Lansmanı", "Pazar", "Ürün", "Plan"], ["2026 Strateji Sunumu", "Öncelikler"]),
        ("egitim_plani.pptx", "Bu sunumun tasarımıyla 'Satış Eğitimi' konulu 3 slaytlık yeni bir sunum hazırla.",
         ["Satış Eğitimi", "Hedefler", "Takvim"], ["Çalışan Eğitim Planı"])):
    add("pptx", "template", [f], f"uploads/{f} — {req}", "create",
        [out("pptx"), {"t": "opens"}, {"t": "pptx_design_from", "template": f, "min_slides": len(titles), "forbid_titles": forbid}],
        {"fn": "pptx_template", "file": f, "titles": titles},
        required=["read_skill", "template", "validate"], forbidden=["bypass"])
add("pptx", "create", [], "professional-blue temasıyla 'Q3 Satış Sonuçları' konulu 5 slaytlık bir sunum hazırla.", "create",
    [out("pptx"), {"t": "opens"}, {"t": "pptx_slides", "count": 5}, {"t": "pptx_bullets_max", "maximum": 6}],
    {"fn": "pptx_create", "name": "q3-sunum.pptx", "titles": ["Q3 Satış Sonuçları", "Özet", "Bölgeler", "Riskler", "Sonraki Adımlar"]},
    required=["read_skill", "create_pptx", "validate"], forbidden=[])
add("pptx", "create", [], "minimal temasıyla 4 slaytlık 'Ekip Tanıtımı' sunumu hazırla.", "create",
    [out("pptx"), {"t": "opens"}, {"t": "pptx_slides", "count": 4}, {"t": "pptx_bullets_max", "maximum": 6}],
    {"fn": "pptx_create", "name": "ekip.pptx", "titles": ["Ekip Tanıtımı", "Biz kimiz", "Ne yapıyoruz", "İletişim"]},
    required=["read_skill", "create_pptx", "validate"], forbidden=[])
add("pptx", "create", [], "Çeyrek bazında ciroyu (Q1 10, Q2 12, Q3 14 milyon TL) gösteren bir sütun grafiği içeren 3 slaytlık bir sunum hazırla.", "create",
    [out("pptx"), {"t": "opens"}, {"t": "pptx_chart", "min_charts": 1}],
    {"fn": "pptx_create", "name": "ciro.pptx", "titles": ["Ciro", "Özet"], "chart": True},
    required=["read_skill", "create_pptx", "validate"], forbidden=[])
add("pptx", "legacy", ["yonetim_sunumu.ppt"], "uploads/yonetim_sunumu.ppt dosyasında 3. slaytın başlığını 'Müşteri tabanı büyüdü' yap.", "edit",
    [out("pptx"), {"t": "pptx_text", "slide": 3, "has": ["Müşteri tabanı büyüdü"], "lacks": ["genişledi"]}],
    {"fn": "legacy", "file": "yonetim_sunumu.ppt", "target": "pptx",
     "then": {"fn": "pptx_ops", "ops": [{"op": "replace", "slide": 3, "old": "Müşteri tabanı genişledi", "new": "Müşteri tabanı büyüdü"}]}},
    required=["read_skill", "legacy", "docgen_pptx"], forbidden=["bypass"],
    judge=J_EDIT + " Dosyanın .pptx biçimine dönüştürüldüğünü belirtmeli.")
add("pptx", "nochange", ["kurumsal_deste.pptx"], "uploads/kurumsal_deste.pptx dosyasında hiçbir şeyi değiştirme, olduğu gibi geri ver.",
    "nochange", [out("pptx"), {"t": "identical"}], {"fn": "copy", "file": "kurumsal_deste.pptx"})

# =============================================================================
# XLSX
# =============================================================================
def sheet_parts(f, *names):
    pkg = Package(FX / f)
    return [X.sheet_part(pkg, n) for n in names]


XWB = ["xl/workbook.xml", "xl/calcChain.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml", "xl/sharedStrings.xml"]
VAL = [
    ("pro_satis_paneli.xlsx", "Satışlar sayfasında 2. satırdaki adedi 16'dan 18'e çıkar.", "Satışlar!E2", 18,
     [("Satışlar!G2", 16200), ("Satışlar!H2", 19440)], ["Satışlar", "Özet"]),
    ("pro_satis_paneli.xlsx", "Parametreler sayfasında KDV oranını %18 yap.", "Parametreler!B2", 0.18, [], ["Parametreler", "Satışlar"]),
    ("fx_fatura.xlsx", "Fatura sayfasında Laptop'un miktarını 2'den 3'e çıkar.", "Fatura!B2", 3, [], ["Fatura"]),
    ("fx_prim.xlsx", "Aslı Kurt'un satış tutarını 120000 yap.", "Primler!B2", 120000, [("Primler!C2", 6000)], ["Primler"]),
    ("satislar.xlsx", "Satışlar sayfasındaki ilk satırın (Marmara, Laptop) adedini 5 yap.", "Satışlar!D2", 5, [("Satışlar!F2", 125000)], ["Satışlar"]),
    ("fx_hatali.xlsx", "Gelir Tablosu'nda Ocak gelirini 420000 yap.", "Gelir Tablosu!B2", 420000, [("Gelir Tablosu!D2", 102000)], ["Gelir Tablosu"]),
]
for f, req, cell, value, cached, sheets in VAL:
    add("xlsx", "value_recalc", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("xlsx"), {"t": "xlsx_value", "cell": cell, "value": value}, {"t": "xlsx_cache_correct"},
         {"t": "xlsx_features_kept"}, {"t": "xlsx_others_same", "except_": [cell]}]
        + [{"t": "xlsx_cached", "cell": c, "value": v} for c, v in cached]
        + [{"t": "parts_only", "allow": sheet_parts(f, *sheets) + XWB}],
        {"fn": "xlsx_ops", "file": f, "ops": [{"op": "set_value", "cell": cell, "value": value}]},
        required=EDIT["xlsx"] + ["recalc"])

add("xlsx", "formula", ["butce.xlsx"], "uploads/butce.xlsx dosyasında Bütçe sayfasındaki Toplam sütununu (F2:F5) sabit değer yerine formülle hesaplat.",
    "edit", [out("xlsx")] + [{"t": "xlsx_formula", "cell": f"Bütçe!F{r}", "contains": "SUM|\\+"} for r in range(2, 6)] +
    [{"t": "xlsx_cached", "cell": "Bütçe!F2", "value": 1660000}, {"t": "xlsx_cache_correct"}],
    {"fn": "xlsx_ops", "file": "butce.xlsx", "ops": [{"op": "set_formula", "cell": f"Bütçe!F{r}", "formula": f"=SUM(B{r}:E{r})"} for r in range(2, 6)],
     "onay": True},
    required=EDIT["xlsx"] + ["recalc"])
add("xlsx", "formula", ["fx_notlar.xlsx"], "uploads/fx_notlar.xlsx dosyasında Ortalama sütununu (F) Vize %30, Final %50, Ödev %20 ağırlıkla formülle hesapla.",
    "edit", [out("xlsx"), {"t": "xlsx_formula", "cell": "Notlar!F2", "contains": "C2"}, {"t": "xlsx_formula", "cell": "Notlar!F31", "contains": "D31"},
             {"t": "xlsx_any_formula", "pattern": ".", "min_count": 30}, {"t": "xlsx_cache_correct"}],
    {"fn": "xlsx_ops", "file": "fx_notlar.xlsx", "ops": [{"op": "set_formula", "cell": f"Notlar!F{r}", "formula": f"=C{r}*0.3+D{r}*0.5+E{r}*0.2"}
                                                       for r in range(2, 32)]},
    required=EDIT["xlsx"] + ["recalc"])
add("xlsx", "formula", ["satislar.xlsx"], "uploads/satislar.xlsx dosyasında Satışlar sayfasının altına Tutar sütununun toplamını veren bir 'Toplam' satırı ekle.",
    "edit", [out("xlsx"), {"t": "xlsx_any_formula", "pattern": "SUM\\(F2:F13\\)", "min_count": 1}, {"t": "xlsx_cache_correct"},
             {"t": "text_has", "texts": ["Toplam"]}],
    {"fn": "xlsx_ops", "file": "satislar.xlsx", "ops": [{"op": "set_value", "cell": "Satışlar!A14", "value": "Toplam"},
                                                      {"op": "set_formula", "cell": "Satışlar!F14", "formula": "=SUM(F2:F13)"}]},
    required=EDIT["xlsx"] + ["recalc"])

XST = [
    ("bicimli.xlsx", "Başlık satırını lacivert dolgu ve beyaz yazı yap.", "Durum!A1:C1", {"font": {"color": "white", "bold": True}, "fill": "navy"},
     {"font": {"color": "FFFFFF"}, "fill": "1F3864"}),
    ("bicimli.xlsx", "Durum sütununu ortala.", "Durum!B1:B4", {"align": "center"}, {"align": {"horizontal": "center"}}),
    ("pro_satis_paneli.xlsx", "Tutar sütununu (G2:G41) binlik ayraçlı ve iki ondalıklı göster.", "Satışlar!G2:G41", {"numfmt": "#,##0\\.00"},
     {"numfmt": "#,##0.00"}),
    ("bicimli.xlsx", "Durumu 'Başarısız' olan hücreleri koyu kırmızı dolguyla işaretle.", "Durum!B3:B4", {"fill": "darkred"}, {"fill": "8B0000"}),
    ("fx_fatura.xlsx", "Fatura sayfasının başlık satırını kalın ve açık gri dolgulu yap.", "Fatura!A1:F1", {"font": {"bold": True}, "fill": "lightgray"},
     {"font": {"bold": True}, "fill": "D9D9D9"}),
    ("fx_prim.xlsx", "Prim sütunundaki değerleri (C2:C14) yeşil yazı rengiyle göster.", "Primler!C2:C14", {"font": {"color": "green"}},
     {"font": {"color": "00B050"}}),
    ("fx_hatali.xlsx", "Gelir sütununu (B2:B13) binlik ayraçlı göster.", "Gelir Tablosu!B2:B13", {"numfmt": "#,##0"}, {"numfmt": "#,##0"}),
    ("satislar.xlsx", "başlık satırının yazı tipini Arial 12 punto yap.", "Satışlar!A1:F1", {"font": {"font": "Arial", "size": 12}},
     {"font": {"name": "Arial", "size": 12}}),
    ("butce.xlsx", "Bütçe sayfasındaki Toplam sütununu (F1:F5) kalın ve sarı dolgulu yap.", "Bütçe!F1:F5", {"font": {"bold": True}, "fill": "yellow"},
     {"font": {"bold": True}, "fill": "FFFF00"}),
    ("fx_notlar.xlsx", "Final sütununun başlığını (D1) italik ve 14 punto yap.", "Notlar!D1:D1", {"font": {"italic": True, "size": 14}},
     {"font": {"italic": True, "size": 14}}),
]
for f, req, rng, want, setp in XST:
    sheet = rng.split("!")[0]
    add("xlsx", "cell_format", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("xlsx"), {"t": "xlsx_style", "range": rng, **want}, {"t": "xlsx_others_same", "except_": [rng]},
         {"t": "xlsx_features_kept"}, {"t": "parts_only", "allow": sheet_parts(f, sheet) + ["xl/styles.xml"] + XWB}],
        {"fn": "xlsx_ops", "file": f, "ops": [{"op": "style", "range": rng, **setp}], "recalc": False})
add("xlsx", "cell_format", ["bicimli.xlsx"], "uploads/bicimli.xlsx dosyasında 'Başarısız' hücrelerindeki altı çizgiyi kaldır; diğer biçimler kalsın.",
    "edit", [out("xlsx"), {"t": "xlsx_style", "range": "Durum!B3:B4", "font": {"underline": False}, "keep": ["bold", "italic", "fill", "color"]},
             {"t": "parts_only", "allow": sheet_parts("bicimli.xlsx", "Durum") + ["xl/styles.xml"] + XWB}],
    {"fn": "xlsx_ops", "file": "bicimli.xlsx", "ops": [{"op": "style", "range": "Durum!B3:B4", "font": {"underline": False}}], "recalc": False})

for f, req, rng, sheet, op in (
        ("pro_satis_paneli.xlsx", "Adet 20'den büyük olan hücreleri yeşil dolguyla göster.", "E2", "Satışlar",
         {"op": "cf", "range": "Satışlar!E2:E41", "operator": "greaterThan", "formula": "20", "fill": "C6EFCE"}),
        ("fx_notlar.xlsx", "Finali 50'nin altında olan notları kırmızı dolguyla göster.", "D2", "Notlar",
         {"op": "cf", "range": "Notlar!D2:D31", "operator": "lessThan", "formula": "50", "fill": "FFC7CE"}),
        ("fx_prim.xlsx", "Satışı 100000 ve üzeri olan temsilcileri yeşil yazıyla vurgula.", "B2", "Primler",
         {"op": "cf", "range": "Primler!B2:B14", "operator": "greaterThanOrEqual", "formula": "100000", "font_color": "00B050"}),
        ("fx_hatali.xlsx", "Kârı 100000'in altında olan ayların Kâr hücrelerini turuncu dolguyla işaretle.", "D2", "Gelir Tablosu",
         {"op": "cf", "range": "Gelir Tablosu!D2:D13", "operator": "lessThan", "formula": "100000", "fill": "FFC000"})):
    add("xlsx", "conditional_format", [f], f"uploads/{f} dosyasında {req} (kalıcı koşullu biçim olarak)", "edit",
        [out("xlsx"), {"t": "xlsx_cf", "sheet": sheet, "sqref_has": rng}, {"t": "xlsx_features_kept", "cf_delta": 1},
         {"t": "xlsx_others_same"}, {"t": "parts_only", "allow": sheet_parts(f, sheet) + ["xl/styles.xml"] + XWB}],
        {"fn": "xlsx_ops", "file": f, "ops": [op], "recalc": False})
for f, sheet, col, width, req in (("fx_fatura.xlsx", "Fatura", "A", 30, "Kalem sütununun (A) genişliğini 30 yap."),
                                  ("satislar.xlsx", "Satışlar", "C", 25, "Ürün sütununun (C) genişliğini 25 yap.")):
    add("xlsx", "column_width", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("xlsx"), {"t": "xlsx_width", "sheet": sheet, "col": col, "minimum": width - 0.5}, {"t": "xlsx_others_same"},
         {"t": "parts_only", "allow": sheet_parts(f, sheet) + XWB}],
        {"fn": "xlsx_ops", "file": f, "ops": [{"op": "width", "sheet": sheet, "col": col, "width": width}], "recalc": False})

STRUCT = [
    ("pro_satis_paneli.xlsx", "Satışlar sayfasında 10. satırın üstüne boş bir satır ekle.",
     [{"t": "xlsx_formula", "cell": "Satışlar!G11", "equals": "=E11*F11"}, {"t": "xlsx_formula", "cell": "Satışlar!G42", "equals": "=E42*F42"},
      {"t": "xlsx_formula", "cell": "Özet!B4", "contains": "\\$42"}, {"t": "xlsx_table_ref", "name": "SatisTablosu", "ref": "A1:H42"},
      {"t": "xlsx_cached", "cell": "Özet!B4", "value": 110050}, {"t": "xlsx_others_same", "shifted": {"sheet": "Satışlar", "at": 10, "n": 1}}],
     [{"op": "insert_rows", "sheet": "Satışlar", "at": 10}]),
    ("pro_satis_paneli.xlsx", "Satışlar tablosunu Tutar'a göre büyükten küçüğe sırala.",
     [{"t": "xlsx_sorted", "sheet": "Satışlar", "col": "G", "first": 2, "last": 41, "desc": True, "keep_rows_from": "ABCDEF"},
      {"t": "xlsx_formula", "cell": "Satışlar!G2", "equals": "=E2*F2"}, {"t": "xlsx_cached", "cell": "Özet!B4", "value": 110050}],
     [{"op": "sort", "sheet": "Satışlar", "range": "A1:H41", "by": "G", "desc": True}]),
    ("satislar.xlsx", "Satışlar tablosunu Bölge'ye göre A'dan Z'ye sırala.",
     [{"t": "xlsx_sorted", "sheet": "Satışlar", "col": "B", "first": 2, "last": 13, "keep_rows_from": "ABCDE"},
      {"t": "xlsx_formula", "cell": "Satışlar!F2", "equals": "=D2*E2"}],
     [{"op": "sort", "sheet": "Satışlar", "range": "A1:F13", "by": "B"}]),
    ("fx_notlar.xlsx", "Notlar tablosunu Final notuna göre büyükten küçüğe sırala.",
     [{"t": "xlsx_sorted", "sheet": "Notlar", "col": "D", "first": 2, "last": 31, "desc": True, "keep_rows_from": "ABCDE"}],
     [{"op": "sort", "sheet": "Notlar", "range": "A1:G31", "by": "D", "desc": True}]),
    ("satislar.xlsx", "Satışlar sayfasında 5. satırı sil.",
     [{"t": "xlsx_formula", "cell": "Satışlar!F5", "equals": "=D5*E5"}, {"t": "xlsx_others_same", "shifted": {"sheet": "Satışlar", "at": 6, "n": -1},
                                                                         "except_": ["Satışlar!A5:F5"]}],
     [{"op": "delete_rows", "sheet": "Satışlar", "at": 5}]),
    ("satislar.xlsx", "Satışlar sayfasında Adet sütununun önüne 'İndirim' başlıklı boş bir sütun ekle.",
     [{"t": "xlsx_value", "cell": "Satışlar!D1", "value": "İndirim"}, {"t": "xlsx_formula", "cell": "Satışlar!G2", "equals": "=E2*F2"}],
     [{"op": "insert_cols", "sheet": "Satışlar", "at": "D"}, {"op": "set_value", "cell": "Satışlar!D1", "value": "İndirim"}]),
    ("fx_prim.xlsx", "Primler sayfasında 2. satırın üstüne boş bir satır ekle.",
     [{"t": "xlsx_formula", "cell": "Primler!C3", "equals": "=IF(B3>=100000,B3*0.05,0)"}],
     [{"op": "insert_rows", "sheet": "Primler", "at": 2}]),
]
for f, req, checks, ops in STRUCT:
    add("xlsx", "structure", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("xlsx"), {"t": "xlsx_features_kept"}, {"t": "xlsx_cache_correct"}] + checks,
        {"fn": "xlsx_ops", "file": f, "ops": ops},
        required=EDIT["xlsx"] + ["xlsx_struct"])

for f, req, why in (
        ("pro_satis_paneli.xlsx", "Parametreler sayfasındaki KDV Oranı satırını sil.", "H sütunundaki formüller ve KDV_Orani adı bu satırı kullanıyor; silmek #REF! hatası doğurur"),
        ("pro_satis_paneli.xlsx", "Satışlar sayfasında G5 hücresine 99000 yaz.", "G5 bir formül içeriyor (=E5*F5); üzerine sabit yazmak hesabı bozar"),
        ("pro_satis_paneli.xlsx", "Satışlar sayfasında E3'teki 24'ü 9 yap.", "E3'te 24 değil 19 var; kullanıcının belirttiği eski değer tutmuyor"),
        ("pro_satis_paneli.xlsx", "Özet sayfasındaki formülleri sabit değerlere çevir.", "formülleri değere çevirmek geri alınamaz bir işlemdir; onay gerekir")):
    add("xlsx", "guard", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge(why) + " Gerekirse onay (--onay) ile yapılabileceğini söyleyebilir.")
add("xlsx", "guard_confirmed", ["pro_satis_paneli.xlsx"], "uploads/pro_satis_paneli.xlsx dosyasında Satışlar sayfasındaki G5 hücresine 99000 yaz. --onay",
    "edit", [out("xlsx"), {"t": "xlsx_value", "cell": "Satışlar!G5", "value": 99000}, {"t": "xlsx_formula", "cell": "Satışlar!G5", "constant": True},
             {"t": "xlsx_cache_correct"}],
    {"fn": "xlsx_ops", "file": "pro_satis_paneli.xlsx", "ops": [{"op": "set_value", "cell": "Satışlar!G5", "value": 99000}], "onay": True},
    forbidden=NO_REBUILD["xlsx"])
for f, req, cands in (("bicimli.xlsx", "Başarısız olan servisin notunu 'acil' yap.", "Servis B ve Servis C"),
                      ("satislar.xlsx", "Marmara satırındaki adedi 4 yap.", "Marmara bölgesindeki birden fazla satır")):
    add("xlsx", "ambiguity", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge("birden fazla satır eşleşiyor") + f" Adayları ({cands}) listeleyip hangisinin kastedildiğini sormalı.")

for f in ("teklif_sablonu.xlsx", "pro_teklif_sablonu.xltx"):
    add("xlsx", "template", [f], f"uploads/{f} — Bu teklif şablonunu doldur: Müşteri 'ACME Ltd', Teklif No 'T-2026-001', Tarih '27.10.2026'; "
        "1. kalem: 'Danışmanlık', miktar 10, birim fiyat 1500.", "create",
        [out("xlsx"), {"t": "text_has", "texts": ["ACME Ltd", "T-2026-001", "Danışmanlık"]}, {"t": "xlsx_cached", "cell": "Teklif!E10", "value": 15000},
         {"t": "xlsx_cached", "cell": "Teklif!E14", "value": 18000}, {"t": "xlsx_cache_correct"}],
        {"fn": "xlsx_template", "file": f, "values": {"Teklif!B5": "ACME Ltd", "Teklif!B6": "T-2026-001", "Teklif!B7": "27.10.2026",
                                                     "Teklif!B10": "Danışmanlık", "Teklif!C10": 10, "Teklif!D10": 1500}},
        required=["read_skill", "docgen_xlsx", "recalc"], forbidden=["bypass"])
add("xlsx", "create", [], "4 gider kalemi ve Q1–Q4 sütunları olan bir bütçe çalışma kitabı oluştur; satır ve sütun toplamları formülle hesaplansın.", "create",
    [out("xlsx"), {"t": "opens"}, {"t": "xlsx_any_formula", "pattern": "SUM", "min_count": 5}, {"t": "xlsx_cache_correct"}],
    {"fn": "xlsx_create", "name": "butce-2026.xlsx", "kind": "budget"},
    required=["read_skill", "create_xlsx", "recalc"], forbidden=[])
add("xlsx", "create", [], "Tarih, Ürün, Adet, Fiyat ve Tutar sütunları olan, Tutar'ı formülle (Adet×Fiyat) hesaplayan, 5 örnek satırlı ve toplam satırı olan bir satış takip tablosu oluştur.",
    "create", [out("xlsx"), {"t": "opens"}, {"t": "xlsx_any_formula", "pattern": "\\*", "min_count": 5},
               {"t": "xlsx_any_formula", "pattern": "SUM", "min_count": 1}, {"t": "xlsx_cache_correct"}],
    {"fn": "xlsx_create", "name": "satis-takip.xlsx", "kind": "sales"},
    required=["read_skill", "create_xlsx", "recalc"], forbidden=[])
add("xlsx", "create", [], "Personel izin takibi için bir Excel hazırla: Personel, Başlangıç, Bitiş, Onaylandı sütunları; Onaylandı sütunu yalnızca 'Evet' veya 'Hayır' kabul etsin.",
    "create", [out("xlsx"), {"t": "opens"}, {"t": "xlsx_dv", "sheet": None, "min_rules": 1}],
    {"fn": "xlsx_create", "name": "izin-takip.xlsx", "kind": "leave"},
    required=["read_skill", "create_xlsx"], forbidden=[])
add("xlsx", "legacy", ["butce.xls"], "uploads/butce.xls dosyasında Personel satırının Q1 değerini 420000 yap.", "edit",
    [out("xlsx"), {"t": "xlsx_value", "cell": "Bütçe!B2", "value": 420000}],
    {"fn": "legacy", "file": "butce.xls", "target": "xlsx", "then": {"fn": "xlsx_ops", "ops": [{"op": "set_value", "cell": "Bütçe!B2", "value": 420000}],
                                                                  "recalc": False}},
    required=["read_skill", "legacy", "docgen_xlsx"], forbidden=["bypass"],
    judge=J_EDIT + " Dosyanın .xlsx biçimine dönüştürüldüğünü belirtmeli.")
add("xlsx", "nochange", ["pro_satis_paneli.xlsx"], "uploads/pro_satis_paneli.xlsx dosyasında hiçbir şeyi değiştirme, olduğu gibi geri ver.",
    "nochange", [out("xlsx"), {"t": "identical"}], {"fn": "copy", "file": "pro_satis_paneli.xlsx"})

# =============================================================================
# PDF
# =============================================================================
INSERT_PARAS = [
    "Şirketin finansal yapısı yıl boyunca istikrarlı kaldı; nakit akışı güçlü seyretti ve borçluluk oranı düştü.",
    "Yatırım bütçesi planlandığı gibi kullanıldı; depo ve filo yatırımları takvimine uygun tamamlandı.",
    "Yönetim maliyet disiplinini sürdürmeyi ve verimlilik projelerine öncelik vermeyi hedefliyor.",
    "Tedarik zincirindeki gecikmeler ikinci yarıda azaldı ve stok devir hızı iyileşti.",
    "Müşteri tahsilat süreleri kısaldı; şüpheli alacaklar önceki yıla göre geriledi.",
    "Kur riskine karşı uygulanan politikalar finansal sonuçlar üzerindeki dalgalanmayı sınırladı.",
]


def pdf_unique(f, text, page=None, near=None):
    return len(F.find(str(FX / f), text, page=page, near=near)) == 1


PREP = [
    ("faaliyet_raporu.pdf", 4, "Rota planlama yazılımı tüm bölgelerde kullanılmaya başlandı.", "Rota planlama yazılımı Mart 2025'ten beri tüm bölgelerde kullanılıyor.", "Operasyon bölümündeki"),
    ("faaliyet_raporu.pdf", 1, "48.600 sevkiyat", "49.100 sevkiyat", "Yönetici Özeti'ndeki"),
    ("faaliyet_raporu.pdf", 5, "Nitelikli sürücü bulmak giderek zorlaşmaktadır.", "Nitelikli sürücü bulmak zorlaşmaya devam etmektedir.", "Riskler bölümündeki"),
    ("faaliyet_raporu.pdf", 2, "En yüksek kâr dördüncü çeyrekte elde edildi.", "En yüksek kâr Q4'te elde edildi.", "Finansal Sonuçlar bölümündeki"),
    ("fiyat_listesi.pdf", 1, "Fiyatlara KDV dahil değildir.", "Fiyatlara %20 KDV dahil değildir.", ""),
    ("fiyat_listesi.pdf", 1, "1 Ocak 2026", "1 Şubat 2026", ""),
    ("calisma_yonergesi.pdf", 1, "temel kuralları belirler", "usul ve esasları belirler", "Amaç bölümündeki"),
    ("personel_listesi.pdf", 1, "Liste her ayın ilk iş günü güncellenir.", "Liste her hafta güncellenir.", ""),
    ("pro_gorselli_rapor.pdf", 1, "Konya deposu Eylül 2025'te faaliyete geçti.", "Konya deposu Ekim 2025'te faaliyete geçti.", ""),
    ("ek_tablolar.pdf", 1, "Ek A — Filo Listesi", "Ek A — Araç Filosu", ""),
    ("satis_raporu.pdf", 1, "Karadeniz bölgesi", "Karadeniz Bölgesi", "Değerlendirme bölümündeki"),
    ("pro_form.pdf", 1, "Izin Talep Formu", "İzin Talep Formu", "başlıktaki"),
]
for f, page, old, new, where in PREP:
    assert pdf_unique(f, old), (f, old)
    pages = len(F.page_texts(str(FX / f)))
    add("pdf", "text_replace", [f], f"uploads/{f} dosyasında {where + ' ' if where else ''}'{old}' ifadesini '{new}' olarak değiştir.", "edit",
        [out("pdf"), {"t": "pdf_pages", "n": pages}, {"t": "pdf_page_has", "page": page, "has": [new], "lacks": [old] if old not in new else []},
         {"t": "pdf_pages_same", "mapping": {str(i): i for i in range(1, pages + 1) if i != page}},
         {"t": "pdf_page_others_same", "page": page, "changed": [old.split()[0], new.split()[0]]},
         {"t": "pdf_font_kept", "page": page, "new_text": new, "like_text": old}],
        {"fn": "pdf_ops", "file": f, "ops": [{"op": "replace", "old": old, "new": new, "page": page}]})

PCELL = [
    ("fiyat_listesi.pdf", 1, "Defter'in birim fiyatını 47,50 yap.", "45,00", "47,50", "Defter"),
    ("fiyat_listesi.pdf", 1, "Makas'ın stok adedini 60 yap.", "75", "60", "Makas"),
    ("faaliyet_raporu.pdf", 3, "Bölgesel Dağılım tablosunda Ege'nin sevkiyat sayısını 9.950 yap.", "9.720", "9.950", "Ege"),
    ("faaliyet_raporu.pdf", 2, "Finansal Sonuçlar tablosunda Q3 kârını 9.300 yap.", "9.100", "9.300", "Q3"),
    ("personel_listesi.pdf", 1, "Can Öztürk'ün unvanını 'Müdür' yap.", "Şef", "Müdür", "Can Öztürk"),
    ("pro_gorselli_rapor.pdf", 2, "Kapasite tablosunda Konya'nın doluluk oranını 74 yap.", "71", "74", "Konya"),
]
for f, page, req, old, new, key in PCELL:
    assert pdf_unique(f, old, near=key), (f, old, key)
    pages = len(F.page_texts(str(FX / f)))
    add("pdf", "table_cell", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("pdf"), {"t": "pdf_pages", "n": pages}, {"t": "pdf_page_has", "page": page, "has": [f"{key}"]},
         {"t": "pdf_pages_same", "mapping": {str(i): i for i in range(1, pages + 1) if i != page}},
         {"t": "pdf_page_others_same", "page": page, "changed": [key]}, {"t": "pdf_row_has", "page": page, "key": key, "value": new}],
        {"fn": "pdf_ops", "file": f, "ops": [{"op": "replace", "old": old, "new": new, "page": page, "near": key}]})

add("pdf", "header_footer", ["faaliyet_raporu.pdf"],
    "uploads/faaliyet_raporu.pdf dosyasında tüm sayfaların üst bilgisini 'Kuzey Lojistik A.Ş. — 2025 Faaliyet Raporu — Gizli' yap.", "edit",
    [out("pdf"), {"t": "pdf_pages", "n": 5}, {"t": "pdf_every_page_has", "text": "Faaliyet Raporu — Gizli"},
     {"t": "pdf_page_others_same", "page": 3, "changed": ["Kuzey Lojistik"]}],
    {"fn": "pdf_ops", "file": "faaliyet_raporu.pdf",
     "ops": [{"op": "replace", "old": "Kuzey Lojistik A.Ş. — 2025 Faaliyet Raporu", "new": "Kuzey Lojistik A.Ş. — 2025 Faaliyet Raporu — Gizli",
              "page": p, "band": "header"} for p in range(1, 6)]})
add("pdf", "header_footer", ["pro_gorselli_rapor.pdf"],
    "uploads/pro_gorselli_rapor.pdf dosyasında iki sayfanın üst bilgisindeki 'Tesis Raporu' ifadesini 'Tesis Raporu (Taslak)' yap.", "edit",
    [out("pdf"), {"t": "pdf_every_page_has", "text": "Tesis Raporu (Taslak)"}, {"t": "pdf_pages", "n": 2}],
    {"fn": "pdf_ops", "file": "pro_gorselli_rapor.pdf",
     "ops": [{"op": "replace", "old": "Kuzey Lojistik A.Ş. — Tesis Raporu", "new": "Kuzey Lojistik A.Ş. — Tesis Raporu (Taslak)", "page": p, "band": "header"}
             for p in (1, 2)]})

PG = [
    ("faaliyet_raporu.pdf", "4. sayfayı sil; sayfa numaraları güncellensin.", 4, {"1": 1, "2": 2, "3": 3},
     [{"t": "pdf_footer", "pattern": "Sayfa {n} / {total}"}, {"t": "pdf_page_has", "page": 4, "has": ["5. Riskler"]}],
     [{"op": "delete", "pages": "4"}, {"op": "renumber"}]),
    ("faaliyet_raporu.pdf", "2. ve 3. sayfaları sil; sayfa numaraları güncellensin.", 3, {"1": 1},
     [{"t": "pdf_footer", "pattern": "Sayfa {n} / {total}"}, {"t": "pdf_page_has", "page": 2, "has": ["4. Operasyon"]}],
     [{"op": "delete", "pages": "2-3"}, {"op": "renumber"}]),
    ("ek_tablolar.pdf", "yalnızca ilk sayfayı bırak, diğerlerini çıkar.", 1, {"1": 1}, [], [{"op": "keep", "pages": "1"}]),
    ("faaliyet_raporu.pdf", "5. sayfayı 4. sayfanın önüne al.", 5, {"1": 1, "2": 2, "3": 3, "4": 5, "5": 4}, [],
     [{"op": "move", "pages": "5", "before": 4}]),
    ("pro_gorselli_rapor.pdf", "2. sayfayı sil.", 1, {"1": 1}, [], [{"op": "delete", "pages": "2"}]),
    ("ek_tablolar.pdf", "2. sayfayı 90 derece döndür.", 2, {"1": 1}, [{"t": "pdf_rotated", "page": 2, "degrees": 90}],
     [{"op": "rotate", "pages": "2", "degrees": 90}]),
]
for f, req, n, mapping, checks, ops in PG:
    add("pdf", "pages", [f], f"uploads/{f} dosyasında {req}", "edit",
        [out("pdf"), {"t": "pdf_pages", "n": n}, {"t": "pdf_pages_same", "mapping": mapping}] + checks,
        {"fn": "pdf_ops", "file": f, "ops": ops}, required=PAGE_STEPS)
add("pdf", "pages", ["faaliyet_raporu.pdf", "ek_tablolar.pdf"],
    "uploads/ek_tablolar.pdf dosyasındaki sayfaları uploads/faaliyet_raporu.pdf raporunun sonuna ekle.", "edit",
    [out("pdf"), {"t": "pdf_pages", "n": 7}, {"t": "pdf_pages_same", "mapping": {str(i): i for i in range(1, 6)}},
     {"t": "pdf_page_has", "page": 6, "has": ["Ek A — Filo Listesi"]}],
    {"fn": "pdf_ops", "file": "faaliyet_raporu.pdf", "ops": [{"op": "insert", "other": "ek_tablolar.pdf"}]}, required=PAGE_STEPS)
add("pdf", "pages", ["faaliyet_raporu.pdf", "ek_tablolar.pdf"],
    "uploads/ek_tablolar.pdf sayfalarını uploads/faaliyet_raporu.pdf raporunun 3. sayfasından sonra araya ekle.", "edit",
    [out("pdf"), {"t": "pdf_pages", "n": 7}, {"t": "pdf_page_has", "page": 4, "has": ["Ek A — Filo Listesi"]},
     {"t": "pdf_pages_same", "mapping": {"1": 1, "2": 2, "3": 3, "6": 4, "7": 5}}],
    {"fn": "pdf_ops", "file": "faaliyet_raporu.pdf", "ops": [{"op": "insert", "other": "ek_tablolar.pdf", "at": 4}]}, required=PAGE_STEPS)

for f, req, cands in (("faaliyet_raporu.pdf", "'iki yeni depo' ifadesini 'üç yeni depo' yap.", "1. sayfa (Yönetici Özeti) ve 4. sayfa (Operasyon)"),
                      ("pro_gorselli_rapor.pdf", "'Bursa deposu' ifadesini 'Bursa lojistik merkezi' yap.", "metin içindeki ve şekil altyazısındaki iki geçiş"),
                      ("makale.pdf", "'Bu bölümde konunun temel kavramları' cümlesini kısalt.", "cümlenin geçtiği çok sayıda paragraf")):
    add("pdf", "ambiguity", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge("ifade birden fazla yerde geçiyor") + f" Adayları ({cands}) belirtip hangisinin kastedildiğini sormalı.")
add("pdf", "text_replace", ["faaliyet_raporu.pdf"],
    "uploads/faaliyet_raporu.pdf dosyasında Operasyon bölümündeki 'iki yeni depo' ifadesini 'üç yeni depo' yap.", "edit",
    [out("pdf"), {"t": "pdf_page_has", "page": 4, "has": ["üç yeni depo"]}, {"t": "pdf_page_has", "page": 1, "has": ["iki yeni depo"]},
     {"t": "pdf_pages_same", "mapping": {"1": 1, "2": 2, "3": 3, "5": 5}}],
    {"fn": "pdf_ops", "file": "faaliyet_raporu.pdf", "ops": [{"op": "replace", "old": "iki yeni depo", "new": "üç yeni depo", "page": 4}]})

for f, req, why in (
        ("faaliyet_raporu.pdf", "Riskler bölümündeki 'Akaryakıt' kelimesini kırmızı ve kalın yap.", "mevcut PDF metninin rengi/kalınlığı yerinde değiştirilemiyor; kaynak belge düzenlenmeli"),
        ("fiyat_listesi.pdf", "Defter satırını kalın yap.", "mevcut PDF metninin kalınlığı yerinde değiştirilemiyor"),
        ("calisma_yonergesi.pdf", "başlığı ortala.", "mevcut PDF metninin hizası yerinde değiştirilemiyor"),
        ("faaliyet_raporu.pdf", "PDF'teki bütün yazıların yazı tipini Georgia yap.", "yazı tipi değişikliği PDF'in yeniden üretilmesini gerektirir; onaysız yapılmaz"),
        ("taranmis_belge.pdf", "Belgenin ilk satırındaki başlığı 'Güncel Yönerge' yap.", "taranmış belgede metin katmanı yok; OCR gerekir"),
        ("taranmis_belge.pdf", "Bu belgedeki tarih alanını 27.10.2026 olarak güncelle.", "taranmış belgede metin katmanı yok; OCR gerekir"),
        ("faaliyet_raporu.pdf", "Riskler bölümünü Operasyon bölümünden önceye taşı.", "Riskler ile Sonuç aynı sayfada; sayfa taşımak Sonuç'u da taşır"),
        ("faaliyet_raporu.pdf", "Finansal Sonuçlar bölümünün sonuna şu on iki paragrafı ekle: " +
         " ".join(f"({chr(97 + i)}) {INSERT_PARAS[i % 6]} Bu konu yönetim kurulunda ayrıca değerlendirilecek ve izleme raporlarına eklenecektir."
                  for i in range(12)),
         "metin sayfaya sığmıyor; yeni sayfaların nereye ekleneceği sorulmalı")):
    add("pdf", "refuse", [f], f"uploads/{f} dosyasında {req}", "refuse", [{"t": "no_output"}],
        judge=refuse_judge(why) + " Alternatif (kaynak dosyayı düzenleme, onay, yer belirtme) önermesi iyi olur.")

INSERT_TXT = " ".join(f"({chr(97 + i)}) {p_}" for i, p_ in enumerate(INSERT_PARAS))
add("pdf", "text_insert", ["faaliyet_raporu.pdf"],
    f"uploads/faaliyet_raporu.pdf dosyasında Finansal Sonuçlar bölümünün sonuna (tablonun altındaki cümleden sonra) şu altı paragrafı ekle: {INSERT_TXT}",
    "edit", [out("pdf"), {"t": "pdf_pages", "n": 5}, {"t": "pdf_page_has", "page": 2, "has": [INSERT_PARAS[0][:40], INSERT_PARAS[5][:40]]},
             {"t": "pdf_pages_same", "mapping": {"1": 1, "3": 3, "4": 4, "5": 5}},
             {"t": "pdf_page_others_same", "page": 2, "changed": []},
             {"t": "pdf_font_kept", "page": 2, "new_text": INSERT_PARAS[0][:30], "like_text": "En yüksek kâr"}],
    {"fn": "pdf_ops", "file": "faaliyet_raporu.pdf", "ops": [{"op": "insert_paras", "after": "En yüksek kâr dördüncü çeyrekte elde edildi.",
                                                            "paragraphs": INSERT_PARAS}]},
    required=["read_skill", "pdf_tools", "diff_check"])

for req, name, checks, ref in (
        ("A4 boyutunda, başlığı lacivert '2026 Kırtasiye Fiyat Listesi' olan, 5 ürünlük fiyat tablosu içeren tek sayfalık bir PDF oluştur.",
         "fiyat.pdf", [{"t": "pdf_pages", "n": 1}, {"t": "pdf_color", "text": "2026 Kırtasiye Fiyat Listesi", "color": "navy"},
                       {"t": "pdf_tables", "min_tables": 1}],
         {"fn": "pdf_create", "name": "fiyat.pdf", "title": "2026 Kırtasiye Fiyat Listesi", "color": "#1F3864", "table": True}),
        ("'Katılım Sertifikası' başlıklı, A4 yatay tek sayfalık bir PDF sertifika hazırla.", "sertifika.pdf",
         [{"t": "pdf_landscape"}, {"t": "text_has", "texts": ["Katılım Sertifikası"]}],
         {"fn": "pdf_create", "name": "sertifika.pdf", "title": "Katılım Sertifikası", "landscape": True}),
        ("'Toplantı Tutanağı' başlıklı; katılımcılar, gündem ve kararlar bölümleri olan bir PDF hazırla.", "tutanak.pdf",
         [{"t": "text_has", "texts": ["Toplantı Tutanağı"]}],
         {"fn": "pdf_create", "name": "tutanak.pdf", "title": "Toplantı Tutanağı — katılımcılar, gündem, kararlar"}),
        ("Başlığı koyu kırmızı (#8B0000) 'Haftalık Rapor' olan, bir tablo içeren PDF rapor oluştur.", "rapor.pdf",
         [{"t": "pdf_color", "text": "Haftalık Rapor", "color": "8B0000"}, {"t": "pdf_tables", "min_tables": 1}],
         {"fn": "pdf_create", "name": "rapor.pdf", "title": "Haftalık Rapor", "color": "#8B0000", "table": True})):
    add("pdf", "create", [], req, "create", [out("pdf"), {"t": "opens"}] + checks, ref,
        required=["read_skill", "create_pdf"], forbidden=[])
for f, text in (("faaliyet_raporu.pdf", "TASLAK"), ("fiyat_listesi.pdf", "GİZLİ")):
    pages = len(F.page_texts(str(FX / f)))
    add("pdf", "watermark", [f], f"uploads/{f} dosyasının tüm sayfalarına '{text}' filigranı ekle.", "edit",
        [out("pdf"), {"t": "pdf_pages", "n": pages}, {"t": "pdf_every_page_has", "text": text}],
        {"fn": "pdf_ops", "file": f, "ops": [{"op": "watermark", "text": text}]},
        required=["read_skill", "watermark"], forbidden=["crop_route"])
add("pdf", "nochange", ["fiyat_listesi.pdf"], "uploads/fiyat_listesi.pdf dosyasında hiçbir şeyi değiştirme, olduğu gibi geri ver.",
    "nochange", [out("pdf"), {"t": "identical"}], {"fn": "copy", "file": "fiyat_listesi.pdf"})


def main():
    path = HERE / "cases.csv"
    cols = ["id", "format", "capability", "fixtures", "request", "expect", "checks", "ref", "required", "forbidden", "judge"]
    with open(path, "w", newline="", encoding="utf-8") as fh:
        wr = csv.DictWriter(fh, fieldnames=cols)
        wr.writeheader()
        for c in CASES:
            wr.writerow(c)
    by = {}
    for c in CASES:
        by.setdefault(c["format"], {}).setdefault(c["capability"], 0)
        by[c["format"]][c["capability"]] += 1
    print(f"{len(CASES)} cases -> {path}")
    for fmt, caps in by.items():
        print(f"  {fmt}: {sum(caps.values())}  " + ", ".join(f"{k}={v}" for k, v in caps.items()))


if __name__ == "__main__":
    main()
