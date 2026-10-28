#!/usr/bin/env python3
"""DOCGEN deck tests — new decks keep every text on its own spot.

1. build_themed_deck.js: one body shape per slide holding all bullets as
   paragraphs, with the theme's line spacing; repair → validate →
   render_check PASS.
2. render_check: bullets added one addText at a time (each a shape on the
   placeholder's spot) FAIL with the slide named; with that deck as
   --baseline the same overlap is only a note.
3. render_check does not count a big number behind a title, a text shadow,
   text on a filled shape or stacked text boxes with their own y.
4. repair_pptx: a bullet made of two runs ("Kalın: " + "devamı") stays one
   paragraph; the second run's <a:pPr> is removed; a second repair is clean.
5. With --fixtures: every fixture deck passes render_check (no false alarm).

Needs node + pptxgenjs and Gotenberg (GOTENBERG_URL): run it in the agent image,
on the docgen network, with the repo mounted.

Usage:
    python tests/deck_tests.py [--fixtures evals/fixtures] [--out DIR]
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
THEME = ROOT / "themes" / "dark-corporate.json"

MASTER = """
const PptxGenJS = require("pptxgenjs");
const deck = new PptxGenJS();
deck.layout = "LAYOUT_WIDE";
deck.defineSlideMaster({ title: "M", background: { color: "0F1B2D" }, objects: [
  { placeholder: { options: { name: "PH_TITLE", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.0 } } },
  { placeholder: { options: { name: "PH_BODY", type: "body", x: 0.6, y: 1.6, w: 12.1, h: 4.9 } } } ] });
const T = { color: "EAF2FA", fontSize: 18, fontFace: "Calibri" };
const items = ["Ay yuvarlaktır", "Roketin burnu üçgendir", "Güneş panelleri kare kare dizilir", "Yıldızların beş ucu vardır"];
function slide(title) {
  const s = deck.addSlide({ masterName: "M" });
  s.addText(title, { placeholder: "PH_TITLE", color: "4FC3F7", fontSize: 32, bold: true });
  return s;
}
"""


# ---- helpers --------------------------------------------------------------
def node(js, out):
    script = out.with_suffix(".js")
    script.write_text(MASTER + js + f'\ndeck.writeFile({{ fileName: {json.dumps(str(out))} }});\n', encoding="utf-8")
    r = subprocess.run(["node", str(script)], capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError(r.stderr.strip()[-300:])


def run(script, *args):
    r = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)], capture_output=True, text=True)
    return r.returncode, r.stdout + r.stderr


def render_check(path, *args):
    code, out = run("render_check.py", path, "--json", *args)
    try:
        return code, json.loads(out)
    except ValueError:
        return code, {"errors": [out.strip()[-300:]], "notes": []}


def body_shapes(path, slide):
    """The slide's body-placeholder shapes: [(run texts, <a:pPr> count, has lnSpc) per paragraph]."""
    x = zipfile.ZipFile(path).read(f"ppt/slides/slide{slide}.xml").decode("utf-8")
    return [[(re.findall(r"<a:t>([^<]*)</a:t>", p), p.count("<a:pPr"), "lnSpc" in p)
             for p in re.findall(r"<a:p>(.*?)</a:p>", sp, re.S)]
            for sp in re.findall(r"<p:sp>.*?</p:sp>", x, re.S) if re.search(r'<p:ph [^>]*type="body"', sp)]


# ---- cases ----------------------------------------------------------------
def themed_deck(tmp, check):
    content = {"title": "Uzaydaki Şekiller", "subtitle": "İlkokul",
               "slides": [{"title": "Daire", "bullets": ["Ay yuvarlaktır", "Güneş daire gibi görünür",
                                                         "Gezegenler neredeyse yuvarlaktır", "Top da yuvarlaktır"]},
                          {"title": "Not", "body": "Madde olmayan tek paragraf"}]}
    cj, out = tmp / "content.json", tmp / "themed.pptx"
    cj.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")
    r = subprocess.run(["node", str(SCRIPTS / "build_themed_deck.js"), str(THEME), str(cj), str(out)],
                       capture_output=True, text=True)
    check("themed: builds", r.returncode == 0, r.stderr.strip()[-200:])
    run("repair_pptx.py", out)
    body = body_shapes(out, 2)
    check("themed: one body shape with 4 paragraphs", len(body) == 1 and len(body[0]) == 4, body)
    check("themed: theme line spacing on every bullet", body and all(lns for _, _, lns in body[0]), body)
    code, _ = run("validate.py", out)
    check("themed: validate PASS", code == 0)
    code, rep = render_check(out)
    check("themed: render_check PASS", code == 0, rep.get("errors"))


def overlap(tmp, check):
    bad = tmp / "per_call.pptx"
    node('slide("Yalnızca başlık");\n'
         'const s = slide("Madde başına addText");\n'
         'items.forEach((t) => s.addText(t, { placeholder: "PH_BODY", bullet: true, ...T }));', bad)
    code, rep = render_check(bad)
    errs = " ".join(rep.get("errors", []))
    check("overlap: per-call bullets FAIL", code == 1 and "text drawn over other text" in errs, rep.get("errors"))
    check("overlap: names slide 2 only", "slide 2:" in errs and "slide 1:" not in errs, rep.get("errors"))
    code, rep = render_check(bad, "--baseline", bad)
    check("overlap: same overlap in --baseline is a note", code == 0 and rep.get("notes"), rep)


def no_false_alarm(tmp, check):
    ok = tmp / "layered.pptx"
    node('let s = slide("Büyük sayı arkada");\n'
         's.addText("01", { x: 0.4, y: 0.1, w: 4, h: 2.4, fontSize: 140, color: "1E3A5F", bold: true });\n'
         's = slide("Gölgeli metin");\n'
         's.addText("Gölgeli bir başlık satırı", { x: 1, y: 2, w: 10, h: 1, fontSize: 32, color: "EAF2FA",'
         ' shadow: { type: "outer", blur: 2, offset: 2, angle: 45, color: "000000", opacity: 0.6 } });\n'
         's = slide("Şekil üstünde etiket");\n'
         's.addShape(deck.ShapeType.ellipse, { x: 1, y: 2, w: 3, h: 3, fill: { color: "4FC3F7" } });\n'
         's.addText("Daire", { x: 1, y: 3.2, w: 3, h: 0.6, align: "center", ...T });\n'
         's = slide("Ayrı kutular, kendi y\'leri");\n'
         'items.forEach((t, i) => s.addText(t, { x: 0.8, y: 1.8 + i * 0.8, w: 10, h: 0.7, bullet: true, ...T }));\n'
         's = slide("Tek addText");\n'
         's.addText(items.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { placeholder: "PH_BODY", ...T });',
         ok)
    code, rep = render_check(ok)
    check("no false alarm: layered/shadow/label/stacked/array", code == 0, rep.get("errors"))


def multirun_bullet(tmp, check):
    out = tmp / "multirun.pptx"
    node('const s = slide("İki parçalı madde");\n'
         's.addText([{ text: "Kalın: ", options: { bullet: true, bold: true } }, { text: "devamı", options: { breakLine: true } },\n'
         '            { text: "İkinci madde", options: { bullet: true, breakLine: true } }],\n'
         '           { placeholder: "PH_BODY", ...T, lineSpacing: 28 });', out)
    code, msg = run("repair_pptx.py", out)
    check("multirun: repair fixes it", code == 0 and "fixed" in msg, msg.strip())
    body = body_shapes(out, 1)
    paras = [(runs, ppr) for runs, ppr, _ in body[0]] if len(body) == 1 else body
    check("multirun: bullet stays one paragraph", paras == [(["Kalın: ", "devamı"], 1), (["İkinci madde"], 1)], paras)
    code, _ = run("validate.py", out)
    check("multirun: validate PASS", code == 0)
    code, msg = run("repair_pptx.py", out)
    check("multirun: second repair is clean", "clean" in msg, msg.strip())


def fixtures(fx, check):
    for deck in sorted(Path(fx).glob("*.pptx")):
        code, rep = render_check(deck)
        check(f"fixture {deck.name}: render_check PASS", code == 0, rep.get("errors"))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fixtures", help="dir with the eval fixtures: their decks must pass render_check")
    ap.add_argument("--out", help="keep the test decks here")
    a = ap.parse_args()
    os.environ.setdefault("NODE_PATH", "/opt/docgen/node_modules")
    tmp = Path(a.out) if a.out else Path(tempfile.mkdtemp(prefix="deck_tests_"))
    tmp.mkdir(parents=True, exist_ok=True)
    results = []

    def check(name, ok, detail=""):
        results.append((name, bool(ok), "" if ok else str(detail)[:300]))

    for case in (themed_deck, overlap, no_false_alarm, multirun_bullet):
        try:
            case(tmp, check)
        except Exception as e:
            check(f"{case.__name__}: ran", False, f"{type(e).__name__}: {e}")
    if a.fixtures:
        fixtures(a.fixtures, check)

    w = max(len(n) for n, _, _ in results)
    for name, ok, detail in results:
        print(f"{name:<{w}}  {'PASS' if ok else 'FAIL'}  {detail}")
    bad = sum(not ok for _, ok, _ in results)
    print(f"\n{len(results) - bad}/{len(results)} passed  (decks: {tmp})")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
