#!/usr/bin/env python3
"""DOCGEN capability tests — the acceptance-checklist scenarios run through
the docgen helpers, each result checked with diff_check (declared scope) and
validate --baseline. Expected refusals must be refusals.

Usage (inside the agent image, or any env with requirements.txt installed):
    python tests/capability_tests.py --fixtures <dir with the eval fixtures> --out <dir>

The fixtures are the acceptance-test files (bicimli.docx, kurumsal_deste.pptx,
pro_satis_paneli.xlsx, faaliyet_raporu.pdf, ...). Recalculation checks need
the formula engine (IronCalc, or LibreOffice); without it they are skipped.
"""
import argparse
import os
import subprocess
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "lib"))
from docgen import Ambiguous, NeedsConfirmation, NotSupported, one  # noqa: E402
from docgen import docx as D, pdf as F, pptx as P, theme as T, xlsx as X  # noqa: E402
from docgen import calc, office  # noqa: E402
from docgen import xlsx_struct as XS  # noqa: E402
from docgen.compare import check  # noqa: E402
from docgen.package import Package  # noqa: E402
from docgen.xlsx import ValueMismatch  # noqa: E402

REFUSALS = (Ambiguous, NeedsConfirmation, NotSupported, ValueMismatch)


def scope(**kw):
    base = {k: [] for k in ("text", "section", "style", "slide", "notes", "page", "sheet", "cell", "feature",
                            "part", "expect", "gone")}
    base.update(kw)
    return base


def first_runs(p):
    return [r.el for r in p.runs]


# (id, fixture, action, scope or "refuse")
DOCX = [
    ("YG-W1", "yapay_zeka_raporu.docx", lambda p: D.replace(one(D.find(p, "önerilir", whole_word=True)), "zorunludur"),
     scope(text=["zorunludur"], text_only=True, expect=["zorunludur"], gone=["önerilir"])),
    ("YG-W2", "pro_kurumsal_rapor.docx", lambda p: D.append_to_section(
        p, one(D.find_paragraphs(p, "Operasyon", headings_only=True)),
        D.new_paragraph("Depo otomasyonu 2026'nın ikinci çeyreğinde devreye alınacak.",
                        like=next(b for b in reversed(D.section_blocks(p, one(D.find_paragraphs(p, "Operasyon", headings_only=True))))
                                  if b.tag == D.w("p") and "".join(b.itertext()).strip()))),
     scope(section=["Operasyon"], expect=["Depo otomasyonu"])),
    ("HD-W1", "satis_raporu.docx", lambda p: (lambda t: D.set_cell_text(t, [r[0] for r in D.table_rows(t)].index("Ege"), 1, "3.250.000"))(one(D.find_table(p, "Ege"))),
     scope(text=["3.250.000"], text_only=True, expect=["3.250.000", "24,9"])),
    ("HD-W2", "satis_raporu.docx", lambda p: D.set_paragraph_text(D.section_blocks(p, one(D.find_paragraphs(p, "Değerlendirme", headings_only=True)))[1],
                                                                  "Marmara ve Ege bölgeleri hedefi aştı; Karadeniz için yeni bir saha planı hazırlanıyor."),
     scope(section=["Değerlendirme"], text_only=True)),
    ("HD-W3", "yapay_zeka_raporu.docx", lambda p: D.delete_section(p, one(D.find_paragraphs(p, "2.1 Veri Kaynakları", headings_only=True))),
     scope(section=["2.1 Veri Kaynakları"], gone=["2.1 Veri Kaynakları"])),
    ("BC-W1", "bicimli.docx", lambda p: D.set_run_props(D.split_runs(one(D.find(p, "kritik sistemler"))), pkg=p, color="006400"),
     scope(text=["kritik sistemler"], format_only=True)),
    ("BC-W2", "bicimli.docx", lambda p: D.set_run_props(D.split_runs(one(D.find(p, "risk analizi", section="Özet"))), pkg=p, font="Times New Roman"),
     scope(section=["Özet"], format_only=True)),
    ("BC-W3", "bicimli.docx", lambda p: D.set_run_props(first_runs(one(D.find_paragraphs(p, "Sonuç olarak dolandırıcılık"))), pkg=p, size=11),
     scope(text=["Sonuç olarak dolandırıcılık"], format_only=True)),
    ("BC-W4", "bicimli.docx", lambda p: D.set_run_props(first_runs(one(D.find_paragraphs(p, "Sonuç olarak dolandırıcılık"))), pkg=p, bold=False),
     scope(text=["Sonuç olarak dolandırıcılık"], format_only=True)),
    ("BC-W5", "basliklar.docx", lambda p: D.set_paragraph_props([one(D.find_paragraphs(p, "Giriş", headings_only=True))], align="center"),
     scope(text=["Giriş"], format_only=True)),
    ("BC-W6", "basliklar.docx", lambda p: [D.set_style_props(p, s, color="7B2D8E") for lvl in range(0, 10) for s in D.style_ids_for_level(p, lvl)],
     scope(style=["Title", "TOCHeading", "KurumsalBaşlık"] + [f"Heading{i}" for i in range(1, 10)])),
    ("CK-W1", "bicimli.docx", lambda p: D.set_run_props(D.split_runs(one(D.find(p, "risk", section="Ek Notlar", whole_word=True))), pkg=p, bold=True),
     scope(section=["Ek Notlar"], format_only=True)),
    ("CK-W2", "bicimli.docx", lambda p: D.set_run_props(D.split_runs(one([m for m in D.find(p, "risk", whole_word=True)
                                                                          if m.para.text[m.end:].startswith(" izleniyor")])), pkg=p, color="0000FF"),
     scope(text=["risk izleniyor"], format_only=True)),
    ("CK-W3", "bicimli.docx", lambda p: D.replace(one(D.find(p, "risk", section="Ek Notlar", whole_word=True)), "tehlike"),
     scope(section=["Ek Notlar"], text_only=True, expect=["risk analizi", "riski", "risk izleniyor"])),
    ("CK-W4", "bicimli.docx", lambda p: D.replace_all(p, "risk", "tehdit", whole_word=True),
     scope(text=["tehdit"], text_only=True, gone=["tehditi"])),
    ("CK-W5", "basliklar.docx", lambda p: one(D.find_paragraphs(p, "Sonuç", headings_only=True), "Sonuç"), "refuse"),
    ("CK-W6", "basliklar.docx", lambda p: D.set_run_props(first_runs(one(D.find_paragraphs(p, "Sonuç", section="Bulgular", headings_only=True))), pkg=p, color="FF0000"),
     scope(section=["Bulgular > Sonuç"], format_only=True)),
    ("CK-W7", "tekrarli_baslik.docx", lambda p: D.delete_section(p, one([h for h in D.headings(p) if h.text.startswith("5.")])),
     scope(section=["5. Sonuç"], expect=["3. Sonuç"], gone=["5. Sonuç"])),
    ("RD-W1", "bicimli.docx", lambda p: D.set_paragraph_props([one(D.find_paragraphs(p, "Önemli uyarı"))], align="left"),
     scope(text=["Önemli uyarı"], format_only=True)),
    ("RD-W2", "bicimli.docx", lambda p: D.set_paragraph_props([one(D.find_paragraphs(p, "Bu paragraf düz metindir"))], line_spacing=1.5),
     scope(section=["Ek Notlar"], format_only=True)),
    ("RD-W3", "yapay_zeka_raporu.docx", lambda p: (lambda t: D.add_column(t, len(D.table_rows(t)[0]) - 1, ["Pay (%)"] + [""] * (len(D.table_rows(t)) - 1)))(one(D.find_table(p, "Kaynak"))),
     scope(section=["Veri Kaynakları"], expect=["Pay (%)"])),
    ("KR-W1", "yapay_zeka_raporu.docx", lambda p: [T.set_color(p, s, c) for s, c in (("accent1", "2E7D32"), ("accent2", "66BB6A"), ("dk2", "1B5E20"))],
     scope(theme=True, format_only=True),
     # LibreOffice draws the cached hex values: none may still follow the old theme
     lambda dst: (lambda n: f"{n} stale theme color caches" if n else None)(
         T.refresh_word_caches(Package(dst), {"accent1", "accent2", "dk2"}))),
]

PPTX = [
    ("YG-P1", "kurumsal_deste.pptx", lambda p: P.replace(one(P.find(p, "Büyüme sürüyor", slide=P.slide(p, title="Sonuç").index)), "Büyüme 2026'da da sürecek"),
     scope(slide=[6], text_only=True)),
    ("HD-P1", "yonetim_sunumu.pptx", lambda p: P.set_notes(p, 3, "Yeni müşteriler çoğunlukla dijital kanaldan geldi."), scope(notes=[3])),
    ("HD-P2", "kurumsal_deste.pptx", lambda p: (lambda t: P.set_cell_text(t, [r[0] for r in P.table_rows(t)].index("Ege"), P.table_rows(t)[0].index("2025"), "5,3"))(one(P.tables(p))),
     scope(slide=[5], text_only=True)),
    ("HD-P3", "yonetim_sunumu.pptx", lambda p: P.delete_slide(p, P.slide(p, title="Teşekkürler").index), scope(structure=True)),
    ("BC-P1", "kurumsal_deste.pptx", lambda p: P.set_run_props(P.runs_of(one(P.find_shapes(p, slide=2, text="Kaynak: İç analiz"))), size=10, italic=True),
     scope(slide=[2], format_only=True)),
    ("BC-P2", "bicimli.pptx", lambda p: P.set_run_props(P.split_runs(one(P.find(p, " seviyede risk var", slide=1))), bold=True), scope(slide=[1], format_only=True)),
    ("BC-P3", "yonetim_sunumu.pptx", lambda p: P.set_run_props(P.runs_of(one(P.find_shapes(p, slide=1, placeholder="title"))), font="Georgia"),
     scope(slide=[1], format_only=True)),
    ("BC-P4", "kurumsal_deste.pptx", lambda p: P.set_fill(one(P.find_shapes(p, slide=2, text="Önemli")), "FF8800"), scope(slide=[2], format_only=True)),
    ("BC-P5", "kurumsal_deste.pptx", lambda p: P.set_background(P.slide(p, title="Sonuç"), "EDEDED"), scope(slide=[6], format_only=True)),
    ("BC-P7", "kurumsal_sunum.pptx", lambda p: T.set_color(p, "accent1", "800020"), scope(theme=True)),
    ("KN-P1", "kurumsal_deste.pptx", lambda p: (lambda sh: P.set_geometry(sh, x=P.slide_size(p)[0] - sh.xfrm[2] - P.cm(1), y=P.cm(1)))(one(P.find_shapes(p, slide=2, text="Önemli"))),
     scope(slide=[2], format_only=True)),
    ("KN-P2", "kurumsal_deste.pptx", lambda p: (lambda sh: P.set_geometry(sh, w=sh.xfrm[2] * 2))(one(P.find_shapes(p, slide=2, text="Kaynak: İç analiz"))),
     scope(slide=[2], format_only=True)),
    ("KN-P3", "kurumsal_deste.pptx", lambda p: P.scale(one(P.find_shapes(p, slide=2, kind="pic")), 0.5), scope(slide=[2], format_only=True)),
    ("CK-P1", "bicimli.pptx", lambda p: P.set_run_props(P.split_runs(one(P.find(p, "risk", slide=2, whole_word=True))), color="FF0000"), scope(slide=[2], format_only=True)),
    ("CK-P2", "bicimli.pptx", lambda p: P.set_run_props(P.split_runs(one(P.find(p, "risk", slide=1, in_title=False, whole_word=True))), underline=True),
     scope(slide=[1], format_only=True)),
    ("CK-P3", "yonetim_sunumu.pptx", lambda p: P.replace(one(P.find(p, "%11", slide=2, in_title=True)), "%12"),
     scope(slide=[2], text_only=True, expect=["%12 arttı", "geçen yıla göre %11"])),
    ("KR-P1", "kurumsal_deste.pptx", lambda p: P.move_slide(p, P.slide(p, title="Süreç").index, P.slide(p, title="Pazar Görünümü").index), scope(structure=True)),
    ("HT-P1", "kurumsal_deste.pptx", lambda p: P.delete_slide(p, P.slide(p, title="Pazar Görünümü").index), "refuse"),
    ("RD-P1", "kurumsal_deste.pptx", lambda p: P.set_paragraph_props(P.paragraphs_of(one(P.find_shapes(p, slide=2, text="Kaynak: İç analiz"))), align="right"),
     scope(slide=[2], format_only=True)),
    ("RD-P2", "kurumsal_deste.pptx", lambda p: (lambda t: P.add_column(t, len(P.table_rows(t)[0]) - 1, ["2026 Hedef"] + [""] * (len(P.table_rows(t)) - 1)))(one(P.tables(p))),
     scope(slide=[5], expect=["2026 Hedef"])),
    ("UY-03", "kurumsal_deste.pptx", lambda p: [P.set_run_props(P.runs_of(sh), size=36) for s in P.slides(p) for sh in P.find_shapes(p, slide=s.index, placeholder="title")],
     scope(slide=[1, 2, 3, 4, 5, 6], format_only=True)),
]

XLSX = [
    ("YG-X1", "pro_satis_paneli.xlsx", lambda p: X.set_value(p, "Satışlar!E2", 18, expect=16), scope(cell=["Satışlar!E2"], recalc=True),
     ["Satışlar!G2=16200", "Satışlar!H2=19440"]),
    ("HD-X1", "pro_satis_paneli.xlsx", lambda p: X.set_value(p, "Parametreler!B2", 0.18, expect=0.2),
     scope(cell=["Parametreler!B2", "Satışlar!H2:H41"], recalc=True), []),
    ("HD-X2", "bicimli.xlsx", lambda p: X.set_value(p, f"Durum!{X.column_of(p, 'Durum', 'Not')}{one(X.rows_where(p, 'Durum', 'Servis', 'Servis B'))}", "yeniden denenecek"),
     scope(text=["yeniden denenecek"]), []),
    ("BC-X1", "bicimli.xlsx", lambda p: X.set_style(p, "Durum!A1:C1", font={"color": "FFFFFF"}, fill="1F3864"), scope(cell=["Durum!A1:C1"]), []),
    ("BC-X2", "pro_satis_paneli.xlsx", lambda p: X.set_style(p, "Satışlar!G2:G41", number_format="#,##0.00"), scope(cell=["Satışlar!G2:G41"]), []),
    ("BC-X3", "bicimli.xlsx", lambda p: X.set_style(p, "Durum!B1:B4", align={"horizontal": "center"}), scope(cell=["Durum!B1:B4"]), []),
    ("BC-X4", "bicimli.xlsx", lambda p: [X.set_style(p, f"Durum!B{r}", font={"underline": False}) for r in X.rows_where(p, "Durum", "Durum", "Başarısız")],
     scope(cell=["Durum!B3:B4"]), []),
    ("BC-X5", "pro_satis_paneli.xlsx", lambda p: X.add_conditional_format(p, "Satışlar!E2:E41", operator="greaterThan", formula="20", fill="C6EFCE"),
     scope(feature=["cf:Satışlar"]), []),
    ("CK-X1", "bicimli.xlsx", lambda p: one(X.rows_where(p, "Durum", "Durum", "Başarısız"), "row"), "refuse", []),
    ("CK-X2", "bicimli.xlsx", lambda p: [X.set_style(p, f"Durum!B{r}", fill="8B0000") for r in X.rows_where(p, "Durum", "Durum", "Başarısız")],
     scope(cell=["Durum!B3:B4"]), []),
    ("HT-X1", "pro_satis_paneli.xlsx", lambda p: XS.delete_rows(p, "Parametreler", 2), "refuse", []),
    ("HT-X2", "pro_satis_paneli.xlsx", lambda p: X.set_value(p, "Satışlar!G5", 99000), "refuse", []),
    ("HT-X3", "pro_satis_paneli.xlsx", lambda p: X.set_value(p, "Satışlar!E3", 9, expect=X.get(p, "Satışlar!E3")[0] + 5), "refuse", []),
    ("KR-X1", "pro_satis_paneli.xlsx", lambda p: XS.insert_rows(p, "Satışlar", 10),
     scope(sheet=["Satışlar"], recalc=True, structure=True, feature=["tables:Satışlar"]), ["Özet!B4=110050"]),
    ("KR-X2", "pro_satis_paneli.xlsx", lambda p: XS.sort_range(p, "Satışlar", "A1:H41", "G", descending=True),
     scope(sheet=["Satışlar"], recalc=True), ["Özet!B4=110050"]),
    ("UY-06", "butce.xlsx", lambda p: (X.set_value(p, "Bütçe!A7", "Notlar"), X.set_value(p, "Bütçe!A8", "Rakamlar taslaktır.")),
     scope(cell=["Bütçe!A7:A8"]), []),
    ("HT-G3", "pro_satis_paneli.xlsx", lambda p: None, scope(), []),
]

PDF_T = [
    ("YG-F1", "faaliyet_raporu.pdf", lambda s, d: F.replace_text(s, d, "Rota planlama yazılımı tüm bölgelerde kullanılmaya başlandı.",
                                                                  "Rota planlama yazılımı Mart 2025'ten beri tüm bölgelerde kullanılıyor."),
     scope(page=[4], expect=["Mart 2025'ten beri"])),
    ("HD-F1", "fiyat_listesi.pdf", lambda s, d: F.replace_text(s, d, "45,00", "47,50", near="Defter"), scope(page=[1], expect=["47,50"])),
    ("HD-F2", "faaliyet_raporu.pdf", lambda s, d: (F.delete_pages(s, d, "4"), F.renumber(d, d)), scope(structure=True, band=True, expect=["Sayfa 4 / 4"])),
    ("KR-F1", "faaliyet_raporu.pdf", lambda s, d: F.insert_pdf(s, d, None), scope(structure=True)),
    ("CK-F1", "faaliyet_raporu.pdf", lambda s, d: F.replace_text(s, d, "iki yeni depo", "üç yeni depo"), "refuse"),
    ("CK-F1b", "faaliyet_raporu.pdf", lambda s, d: F.replace_text(s, d, "iki yeni depo", "üç yeni depo", page=4), scope(page=[4], expect=["iki yeni depo açıldı"])),
    ("UY-01", "faaliyet_raporu.pdf", lambda s, d: [F.replace_text(s if p == 1 else d, d, "Kuzey Lojistik A.Ş. — 2025 Faaliyet Raporu",
                                                                  "Kuzey Lojistik A.Ş. — 2025 Faaliyet Raporu — Gizli", page=p, band="header") for p in range(1, 6)],
     scope(band=True, expect=["Raporu — Gizli"])),
]


def run_validate(result, baseline):
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / "validate.py"), result, "--baseline", baseline],
                       capture_output=True, text=True, env={**os.environ, "DOCGEN_SCHEMAS": str(ROOT / "schemas")})
    return "PASS" if r.returncode == 0 else "FAIL"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fixtures", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    fx, out = Path(a.fixtures), Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    have_calc = calc.available() or office.soffice() is not None
    rows, bad = [], 0

    def report(tid, status, detail=""):
        nonlocal bad
        ok = status.startswith("OK") or status.startswith("REFUSED (expected)")
        bad += 0 if ok else 1
        rows.append((tid, status, detail))

    for tid, fname, action, sc, *extra in DOCX + PPTX + XLSX:
        src = str(fx / fname)
        dst = str(out / f"{tid}{Path(fname).suffix}")
        try:
            pkg = Package(src)
            action(pkg)
            if sc == "refuse":
                report(tid, "UNEXPECTED: no refusal")
                continue
            pkg.save(dst)
            notes = []
            if extra and callable(extra[0]):
                err = extra[0](dst)
                if err:
                    report(tid, "FAIL", err)
                    continue
            if extra and Path(fname).suffix == ".xlsx" and (sc.get("recalc") or extra[0]):
                if have_calc:
                    cmd = [sys.executable, str(ROOT / "scripts" / "xlsx_recalc.py"), dst]
                    for e in extra[0]:
                        cmd += ["--expect", e]
                    r = subprocess.run(cmd, capture_output=True, text=True)
                    notes.append("recalc " + ("PASS" if r.returncode == 0 else "FAIL"))
                    if r.returncode:
                        report(tid, "FAIL recalc", r.stdout.strip().splitlines()[-1])
                        continue
                else:
                    notes.append("recalc skipped (no formula engine)")
            rep = check(src, dst, sc)
            val = run_validate(dst, src)
            detail = "; ".join(notes + [f"validate {val}"])
            if rep["ok"] and val == "PASS":
                report(tid, "OK", detail)
            else:
                why = [f"{e['where']}: {e['what']}" for e in rep["violations"][:2]] + rep["part_violations"][:1] + \
                      rep["expectation_failures"][:1]
                report(tid, "FAIL", f"{detail}; {why}")
        except REFUSALS as e:
            report(tid, "REFUSED (expected)" if sc == "refuse" else "UNEXPECTED refusal", f"{type(e).__name__}: {str(e)[:90]}")
        except Exception:
            report(tid, "ERROR", traceback.format_exc().strip().splitlines()[-1])

    for tid, fname, action, sc in PDF_T:
        src, dst = str(fx / fname), str(out / f"{tid}.pdf")
        try:
            if tid == "KR-F1":
                F.insert_pdf(src, dst, str(fx / "ek_tablolar.pdf"))
            else:
                action(src, dst)
            if sc == "refuse":
                report(tid, "UNEXPECTED: no refusal")
                continue
            rep = check(src, dst, sc)
            report(tid, "OK" if rep["ok"] else "FAIL",
                   "" if rep["ok"] else str([e["where"] for e in rep["violations"][:2]] + rep["expectation_failures"][:1]))
        except REFUSALS as e:
            report(tid, "REFUSED (expected)" if sc == "refuse" else "UNEXPECTED refusal", f"{type(e).__name__}: {str(e)[:90]}")
        except Exception:
            report(tid, "ERROR", traceback.format_exc().strip().splitlines()[-1])

    w = max(len(r[0]) for r in rows)
    for tid, status, detail in rows:
        print(f"{tid:<{w}}  {status:<20}  {detail}")
    print(f"\n{len(rows) - bad}/{len(rows)} as expected" + ("" if have_calc else "  (no formula engine: recalc skipped)"))
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
