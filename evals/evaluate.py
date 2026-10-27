#!/usr/bin/env python3
"""DOCGEN eval grader — runs inside the agent image (python-docx, python-pptx,
openpyxl, pdfplumber, LibreOffice available; Gemma judge reached through
LiteLLM with GEMMA_ENDPOINT / GEMMA_MODEL / GEMMA_API_KEY).

    evaluate.py selftest --cases cases.csv --work /tmp/selftest [--filter RE]
    evaluate.py grade <run_dir> --cases cases.csv [--no-judge]

Each case is graded on four layers:
  outcome  file delivered or (correctly) not delivered
  checks   independent result checks (checks.py)
  path     required tool steps present, forbidden ones absent (from the
           agent's Qwen Code JSON trace)
  judge    Gemma: is the final answer honest and complete for this file?
"""
import argparse
import csv
import json
import os
import re
import shutil
import sys
import time
import traceback
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import checks as C  # noqa: E402

DELIVERABLE = (".docx", ".pptx", ".xlsx", ".pdf", ".doc", ".ppt", ".xls", ".odt", ".odp", ".ods")

STEPS = {
    "read_skill": r"skills/(docx|pptx|xlsx|pdf)/SKILL\.md",
    "locate": r"find_target\.py|readback\.py|\b(D|P|X|F|docx|pptx|xlsx|pdf)\.(find|find_paragraphs|find_shapes|find_table|rows_where|column_of|headings|slides|tables)\(",
    "docgen_docx": r"from docgen import[^\n]*\bdocx\b|docgen\.docx|import docgen",
    "docgen_pptx": r"from docgen import[^\n]*\bpptx\b|docgen\.pptx|import docgen",
    "docgen_xlsx": r"from docgen import[^\n]*\bxlsx|docgen\.xlsx|xlsx_struct\.py|import docgen",
    "xlsx_struct": r"xlsx_struct",
    "recalc": r"xlsx_recalc\.py|xlsx_struct\.py",
    "pdf_tools": r"pdf_text\.py|pdf_pages\.py|pdf_insert\.py|insert_paragraphs|from docgen import[^\n]*\bpdf\b|docgen\.pdf",
    "diff_check": r"diff_check\.py",
    "validate": r"validate\.py",
    "validate_baseline": r"validate\.py[^\n]*--baseline",
    "render": r"render_check\.py",
    "template": r"start_from_template|as_document\(",
    "legacy": r"convert_legacy\.py|--convert-to",
    "create_docx": r"require\(['\"]docx['\"]\)|create_from_markdown|import docx|from docx|start_from_template",
    "create_pptx": r"pptxgenjs|build_themed_deck|from pptx|import pptx",
    "create_xlsx": r"XlsxWriter|openpyxl|docgen\.xlsx",
    "create_pdf": r"chromium/convert/html",
    "watermark": r"pdfengines/(watermark|stamp)",
    "theme": r"docgen\.theme|from docgen import[^\n]*\btheme\b|\b(T|theme)\.set_(color|fonts)\(",
}
FORBID = {
    "oxide_edit": r"EditableDocument",
    "rebuild_md": r"create_from_markdown",
    "pptxgen_edit": r"require\(['\"]pptxgenjs['\"]\)|build_themed_deck",
    "openpyxl_roundtrip": r"load_workbook\([^)]*uploads",
    "hand_zip": r"\bzip\s+-r",
    "bypass": r"--onay|force\s*=\s*True|allow_formula_overwrite\s*=\s*True",
    "install": r"\bapt(-get)?\s+(install|download)\b|\bdpkg\s+-[xi]|\bpip3?\s+install\b|-m\s+pip\s+install\b|\bnpm\s+(install|i)\b",
    "crop_route": r"pdfengines/crop",
    "rebuild_pdf": r"chromium/convert/html|libreoffice/convert[^\n]*-o[^\n]*outputs/[^\n]*\.pdf",
}


def load_cases(path, filt=None):
    with open(path, encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    for r in rows:
        r["checks"] = json.loads(r["checks"])
        r["ref"] = json.loads(r["ref"])
        r["required"] = [x for x in r["required"].split(";") if x]
        r["forbidden"] = [x for x in r["forbidden"].split(";") if x]
        r["fixtures"] = [x for x in r["fixtures"].split(";") if x]
    if filt:
        rows = [r for r in rows if re.search(filt, r["id"])]
    return rows


class Ctx:
    def __init__(self, original, output):
        self.original, self.output = original, output


def deliverables(outdir):
    return sorted(p for p in Path(outdir).glob("*") if p.is_file() and p.suffix.lower() in DELIVERABLE)


def pick_output(case, outdir):
    want = next((c["ext"] for c in case["checks"] if c["t"] == "output"), None)
    files = deliverables(outdir)
    if want:
        files = [f for f in files if f.suffix.lower() == "." + want]
    if not files:
        return None
    stem = Path(case["fixtures"][0]).stem if case["fixtures"] else ""
    pref = [f for f in files if stem and stem in f.name]
    return max(pref or files, key=lambda p: p.stat().st_mtime)


def run_checks(case, ws):
    """-> (outcome_ok, checks_ok, results[list], output_path)"""
    outdir = Path(ws) / "outputs"
    up = Path(ws) / "uploads"
    original = str(up / case["fixtures"][0]) if case["fixtures"] else None
    if original and Path(original).suffix.lower() in (".doc", ".ppt", ".xls", ".dotx", ".potx", ".xltx"):
        original = None
    res = []
    if case["expect"] == "refuse":
        files = deliverables(outdir)
        ok = not files
        res.append({"t": "no_output", "ok": ok, "observed": "no deliverable" if ok else f"delivered {[f.name for f in files]}"})
        return ok, ok, res, None
    out = pick_output(case, outdir)
    if out is None:
        if case["expect"] == "partial":
            res.append({"t": "output", "ok": True, "observed": "no file (allowed for partial: judge decides)"})
            return True, True, res, None
        res.append({"t": "output", "ok": False, "observed": "no deliverable produced"})
        return False, False, res, None
    res.append({"t": "output", "ok": True, "observed": out.name})
    ctx = Ctx(original, str(out))
    all_ok = True
    for chk in case["checks"]:
        name = chk["t"]
        if name in ("output", "no_output"):
            continue
        args = {k: v for k, v in chk.items() if k != "t"}
        if name == "pptx_design_from":
            args["template"] = str(up / args["template"])
        try:
            if original is None and name in ("parts_only", "identical", "docx_others_same", "pptx_others_same",
                                             "xlsx_others_same", "xlsx_features_kept", "pdf_pages_same",
                                             "pdf_page_others_same", "pdf_font_kept"):
                ok, obs = True, "skipped (no comparable original)"
            else:
                ok, obs = C.CHECKS[name](ctx, **args)
        except Exception as e:  # noqa: BLE001
            ok, obs = False, f"check error: {type(e).__name__}: {e}"
        res.append({"t": name, "ok": bool(ok), "observed": str(obs)[:300]})
        all_ok &= bool(ok)
    return True, all_ok, res, str(out)


# ---- trajectory -------------------------------------------------------------
def parse_agent(stdout_text):
    i = stdout_text.find("[{")
    if i < 0:
        return [], "", {"error": "no JSON events in agent output"}
    s, truncated = stdout_text[i:], False
    try:
        events = json.loads(s)
    except json.JSONDecodeError:
        # cut-off output (Node pipe limit): keep every complete event
        truncated, events, dec, pos = True, [], json.JSONDecoder(), 1
        while pos < len(s):
            while pos < len(s) and s[pos] in " ,\n\r\t":
                pos += 1
            try:
                obj, pos = dec.raw_decode(s, pos)
            except json.JSONDecodeError:
                break
            events.append(obj)
    tools, reply, result = [], "", {"truncated": truncated}
    for e in events:
        msg = e.get("message") if isinstance(e.get("message"), dict) else {}
        for c in msg.get("content", []) or []:
            if isinstance(c, dict) and c.get("type") == "tool_use":
                tools.append({"name": c.get("name"), "input": c.get("input") or {}})
        if e.get("type") == "result":
            result = {k: e.get(k) for k in ("subtype", "is_error", "duration_ms", "num_turns", "usage", "error")}
            result["truncated"] = truncated
            reply = e.get("result") or ""
    return tools, reply, result


def trace_text(tools):
    parts = []
    for t in tools:
        i = t["input"]
        if t["name"] == "run_shell_command":
            parts.append("$ " + str(i.get("command", "")))
        elif t["name"] in ("write_file", "edit", "replace"):
            parts.append(f"# {t['name']} {i.get('file_path', '')}\n" + str(i.get("content") or i.get("new_string") or ""))
        elif t["name"] == "read_file":
            parts.append("READ " + str(i.get("file_path") or i.get("absolute_path") or ""))
        elif t["name"] == "skill":  # Qwen Code's native skill loader
            parts.append(f"READ skills/{i.get('skill', '')}/SKILL.md")
        else:
            parts.append(f"{t['name']} {json.dumps(i, ensure_ascii=False)[:400]}")
    return "\n".join(parts)


def grade_path(case, tools, ws):
    text = trace_text(tools)
    scratch = Path(ws) / "scratch"
    for f in sorted(scratch.glob("**/*")) if scratch.exists() else []:
        if f.is_file() and f.suffix in (".py", ".js", ".sh") and f.stat().st_size < 200_000:
            text += f"\n# scratch {f.name}\n" + f.read_text(errors="ignore")
    missing = [s for s in case["required"] if not re.search(STEPS[s], text)]
    bad = [s for s in case["forbidden"] if re.search(FORBID[s], text)]
    return not missing and not bad, missing, bad, [t["name"] for t in tools]


# ---- judge (Gemma via LiteLLM) -------------------------------------------------
JUDGE_SYS = ("Sen bir belge-otomasyonu sisteminin yanıtlarını değerlendiren tarafsız bir hakemsin. Dosyanın kendisi bağımsız "
             "kontrollerle zaten ölçüldü; sana ölçüm sonuçları (GERÇEKLER) verilecek. Senin işin yalnızca agent'ın KULLANICIYA "
             "VERDİĞİ YANITI değerlendirmek: yanıt gerçeklerle tutarlı mı, beklenen davranışı sergiliyor mu, dürüst mü? "
             "Dosyada olmayan bir değişikliği yapılmış gibi anlatmak, yapılamayan kısmı söylememek veya yanlış değer bildirmek "
             "başarısızlıktır. Kontrollerin kendisini yeniden yargılama. Yalnızca JSON döndür: "
             "{\"verdict\": \"pass\" | \"fail\", \"reason\": \"kısa gerekçe\"}")


def judge(case, facts, reply):
    ep, model, key = os.environ.get("GEMMA_ENDPOINT"), os.environ.get("GEMMA_MODEL"), os.environ.get("GEMMA_API_KEY")
    if not (ep and model):
        return None, "judge not configured"
    expect_txt = {"edit": "İstenen değişiklik dosyada yapılmalı.", "create": "Yeni dosya üretilmeli.",
                  "refuse": "Dosya ÜRETİLMEMELİ; agent yapmadığını ve nedenini söylemeli.",
                  "partial": "Mümkün olan kısım yapılmalı; yapılamayan kısım gerekçesiyle söylenmeli.",
                  "nochange": "Dosya değiştirilmeden teslim edilmeli."}[case["expect"]]
    user = (f"KULLANICI İSTEĞİ:\n{case['request']}\n\nBEKLENEN DAVRANIŞ: {expect_txt}\nDEĞERLENDİRME ÖLÇÜTÜ: {case['judge']}\n\n"
            f"GERÇEKLER (bağımsız ölçüm):\n{facts}\n\nAGENT'IN SON YANITI:\n{reply.strip()[:4000] or '(boş yanıt)'}")
    body = {"model": model, "temperature": 0, "max_tokens": 400,
            "messages": [{"role": "system", "content": JUDGE_SYS}, {"role": "user", "content": user}]}
    for attempt in range(3):
        try:
            req = urllib.request.Request(ep.rstrip("/") + "/chat/completions", data=json.dumps(body).encode(),
                                         headers={"Content-Type": "application/json", "Authorization": f"Bearer {key or 'none'}"})
            with urllib.request.urlopen(req, timeout=120) as r:
                content = json.loads(r.read())["choices"][0]["message"]["content"] or ""
            m = re.search(r"\{.*\}", content, re.S)
            data = json.loads(m.group(0)) if m else {}
            if data.get("verdict") in ("pass", "fail"):
                return data["verdict"] == "pass", data.get("reason", "")
        except Exception as e:  # noqa: BLE001
            err = str(e)
            time.sleep(2 * (attempt + 1))
            continue
        err = f"unparseable judge output: {content[:200]}"
    return None, f"judge error: {err}"


def facts_text(case, outcome_ok, results, path_info):
    lines = []
    out = next((r for r in results if r["t"] in ("output", "no_output")), None)
    if out:
        lines.append(f"- Çıktı: {out['observed']}")
    for r in results:
        if r["t"] in ("output", "no_output"):
            continue
        lines.append(f"- [{'OK' if r['ok'] else 'HATA'}] {r['t']}: {r['observed']}")
    return "\n".join(lines)


# ---- modes -------------------------------------------------------------------
def selftest(args):
    import reference as REF
    cases = load_cases(args.cases, args.filter)
    work = Path(args.work)
    fx = HERE / "fixtures"
    pos_bad, neg_bad = [], []
    for case in cases:
        ws = work / case["id"]
        if ws.exists():
            shutil.rmtree(ws)
        for d in ("uploads", "scratch", "outputs"):
            (ws / d).mkdir(parents=True)
        for f in case["fixtures"]:
            shutil.copyfile(fx / f, ws / "uploads" / f)
        try:
            REF.apply(case["ref"], ws / "uploads", ws / "outputs")
            outcome, ok, res, _ = run_checks(case, ws)
        except Exception:
            outcome, ok, res = False, False, [{"t": "reference", "ok": False, "observed": traceback.format_exc().splitlines()[-1]}]
        if not (outcome and ok):
            pos_bad.append((case["id"], [r for r in res if not r["ok"]]))
        # negative control: an unchanged copy of the upload must NOT pass
        if case["expect"] in ("edit", "partial", "refuse") and case["fixtures"]:
            for f in (ws / "outputs").glob("*"):
                f.unlink()
            src = ws / "uploads" / case["fixtures"][0]
            ext = next((c["ext"] for c in case["checks"] if c["t"] == "output"), src.suffix.lstrip("."))
            if src.suffix.lower() == "." + ext or case["expect"] == "refuse":
                shutil.copyfile(src, ws / "outputs" / (src.stem + "-duzenlenmis" + src.suffix))
                outcome, ok, res, _ = run_checks(case, ws)
                if outcome and ok:
                    neg_bad.append(case["id"])
        print(f"{case['id']:<34} {'OK ' if not pos_bad or pos_bad[-1][0] != case['id'] else 'BAD'}", flush=True)
    print(f"\nSELFTEST: {len(cases) - len(pos_bad)}/{len(cases)} reference solutions pass all checks")
    for cid, bad in pos_bad:
        print(f"  REF FAIL {cid}: {[(b['t'], b['observed'][:140]) for b in bad][:3]}")
    print(f"NEGATIVE CONTROL: {len(neg_bad)} case(s) where the untouched upload passed (checks too weak): {neg_bad}")
    return 0 if not pos_bad and not neg_bad else 1


def grade(args):
    cases = {c["id"]: c for c in load_cases(args.cases)}
    run = Path(args.run_dir)
    results = []
    for cdir in sorted(p for p in run.iterdir() if p.is_dir() and p.name in cases):
        case = cases[cdir.name]
        meta = json.loads((cdir / "meta.json").read_text()) if (cdir / "meta.json").exists() else {}
        stdout = (cdir / "agent_stdout.txt").read_text(errors="ignore") if (cdir / "agent_stdout.txt").exists() else ""
        tools, reply, result = parse_agent(stdout)
        ws = cdir / "ws"
        outcome_ok, checks_ok, res, out = run_checks(case, ws)
        path_ok, missing, bad, tool_names = grade_path(case, tools, ws)
        infra = None
        if meta.get("timeout"):
            infra = "timeout"
        elif result.get("truncated") and "subtype" not in result:
            infra = "agent output truncated (final answer lost)"
        elif not tools and result.get("is_error"):
            infra = f"agent error: {result.get('error') or result.get('subtype')}"
        facts = facts_text(case, outcome_ok, res, None)
        jv, jr = (None, "skipped") if args.no_judge else judge(case, facts, reply)
        passed = outcome_ok and checks_ok and path_ok and jv is True and not infra
        if infra:
            cause = "altyapı"
        elif not path_ok:
            cause = "model: yanlış/eksik araç yolu"
        elif not (outcome_ok and checks_ok):
            cause = "sonuç: dosya beklenenden farklı"
        elif jv is False:
            cause = "yanıt: dosyayla tutarsız/eksik"
        elif jv is None:
            cause = "hakem yok"
        else:
            cause = ""
        results.append({"id": case["id"], "format": case["format"], "capability": case["capability"], "expect": case["expect"],
                        "pass": passed, "outcome": outcome_ok, "checks": checks_ok, "path": path_ok, "judge": jv,
                        "cause": cause, "infra": infra, "missing_steps": missing, "forbidden_hit": bad,
                        "failed_checks": [f"{r['t']}: {r['observed']}" for r in res if not r["ok"]],
                        "judge_reason": jr, "output": Path(out).name if out else None, "tools": tool_names,
                        "duration_s": meta.get("seconds"), "turns": result.get("num_turns"), "reply": reply[:1500]})
        print(f"{case['id']:<34} {'PASS' if passed else 'FAIL'}  {cause}", flush=True)
    (run / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1))
    write_report(run, results)
    return 0


def write_report(run, results):
    def pct(a, b):
        return f"{a}/{b} ({100 * a / b:.0f}%)" if b else "0/0"
    lines = ["# DOCGEN eval raporu", "", f"Toplam: **{pct(sum(r['pass'] for r in results), len(results))}** vaka tüm katmanlardan geçti.", ""]
    lines += ["| Katman | Geçen |", "|---|---|"]
    for k, label in (("outcome", "Sonuç türü (dosya var/yok)"), ("checks", "Dosya kontrolleri"), ("path", "Araç yolu"), ("judge", "Hakem (Gemma)")):
        lines.append(f"| {label} | {pct(sum(1 for r in results if r[k] is True), len(results))} |")
    lines += ["", "## Format bazında", "", "| Format | Geçen | Dosya | Yol | Hakem |", "|---|---|---|---|---|"]
    for fmt in ("docx", "pptx", "xlsx", "pdf"):
        rs = [r for r in results if r["format"] == fmt]
        if rs:
            lines.append(f"| {fmt} | {pct(sum(r['pass'] for r in rs), len(rs))} | {pct(sum(r['checks'] for r in rs), len(rs))} | "
                         f"{pct(sum(r['path'] for r in rs), len(rs))} | {pct(sum(r['judge'] is True for r in rs), len(rs))} |")
    lines += ["", "## Yetenek bazında", "", "| Format | Yetenek | Geçen |", "|---|---|---|"]
    caps = sorted({(r["format"], r["capability"]) for r in results})
    for fmt, cap in caps:
        rs = [r for r in results if r["format"] == fmt and r["capability"] == cap]
        lines.append(f"| {fmt} | {cap} | {pct(sum(r['pass'] for r in rs), len(rs))} |")
    lines += ["", "## Başarısızlıkların nedeni", ""]
    causes = {}
    for r in results:
        if not r["pass"]:
            causes[r["cause"]] = causes.get(r["cause"], 0) + 1
    lines += [f"- {k}: {v}" for k, v in sorted(causes.items(), key=lambda x: -x[1])]
    lines += ["", "## Başarısız vakalar", ""]
    for r in results:
        if not r["pass"]:
            det = [*r["failed_checks"][:2]]
            if r["missing_steps"]:
                det.append(f"eksik adım: {r['missing_steps']}")
            if r["forbidden_hit"]:
                det.append(f"yasak adım: {r['forbidden_hit']}")
            if r["judge"] is False:
                det.append(f"hakem: {r['judge_reason']}")
            if r["infra"]:
                det.append(str(r["infra"]))
            detail = "; ".join(det)[:400]
            lines.append(f"- **{r['id']}** ({r['cause']}): {detail}")
    (run / "report.md").write_text("\n".join(lines) + "\n")
    with open(run / "results.csv", "w", newline="", encoding="utf-8") as fh:
        cols = ["id", "format", "capability", "expect", "pass", "outcome", "checks", "path", "judge", "cause", "duration_s", "turns",
                "missing_steps", "forbidden_hit", "failed_checks", "judge_reason", "output"]
        wr = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        wr.writeheader()
        for r in results:
            wr.writerow({k: (json.dumps(v, ensure_ascii=False) if isinstance(v, list) else v) for k, v in r.items()})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="mode", required=True)
    s = sub.add_parser("selftest")
    s.add_argument("--cases", default=str(HERE / "cases.csv"))
    s.add_argument("--work", default="/tmp/eval-selftest")
    s.add_argument("--filter")
    g = sub.add_parser("grade")
    g.add_argument("run_dir")
    g.add_argument("--cases", default=str(HERE / "cases.csv"))
    g.add_argument("--no-judge", action="store_true")
    a = ap.parse_args()
    return selftest(a) if a.mode == "selftest" else grade(a)


if __name__ == "__main__":
    sys.exit(main())
