#!/usr/bin/env python3
"""DOCGEN eval runner (host side, stdlib only).

Runs every case through the REAL system — same image, same entrypoint, same
`docgen` network with Gotenberg, same .env (Qwen as the agent model) — one
throw-away container per case, then grades the run inside the image with
evaluate.py (independent file checks, tool path, Gemma judge).

    python3 evals/run_eval.py                       # all cases, 4 in parallel
    python3 evals/run_eval.py --filter 'DOCX-run_format' --concurrency 6
    python3 evals/run_eval.py --grade-only evals/runs/<run>   # re-grade
    python3 evals/run_eval.py --selftest            # prove the checks first

Results: evals/runs/<timestamp>/{report.md,results.csv,results.json} and one
folder per case with the agent's workspace and raw output.
"""
import argparse
import concurrent.futures as cf
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EVALS = ROOT / "evals"
IMAGE = "docgen-agent:latest"


def sh(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def load_cases(path, filt, limit):
    with open(path, encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    if filt:
        rows = [r for r in rows if re.search(filt, r["id"])]
    return rows[:limit] if limit else rows


def ensure_gotenberg():
    sh(["docker", "compose", "up", "-d", "gotenberg"], cwd=ROOT)
    for _ in range(30):
        r = sh(["docker", "run", "--rm", "--network", "docgen", "--entrypoint", "curl", IMAGE, "-sf",
                "http://gotenberg:3000/health"])
        if r.returncode == 0:
            return True
        time.sleep(2)
    return False


def run_case(case, run_dir, timeout):
    cdir = run_dir / case["id"]
    ws = cdir / "ws"
    if cdir.exists():
        shutil.rmtree(cdir)
    for d in ("uploads", "scratch", "outputs"):
        (ws / d).mkdir(parents=True)
    for f in [x for x in case["fixtures"].split(";") if x]:
        shutil.copyfile(EVALS / "fixtures" / f, ws / "uploads" / f)
    os.system(f"chmod -R a+rwX '{cdir}'")
    name = f"docgen-eval-{case['id'].lower().replace('_', '-')}"
    cmd = ["docker", "run", "--rm", "--name", name, "--network", "docgen", "--env-file", str(ROOT / ".env"),
           "-e", "GOTENBERG_URL=http://gotenberg:3000",
           "-v", f"{ws}:/workspace", "-v", f"{ws / 'uploads'}:/workspace/uploads:ro", IMAGE, case["request"]]
    start = time.time()
    meta = {"timeout": False}
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        out, err, meta["exit"] = r.stdout, r.stderr, r.returncode
    except subprocess.TimeoutExpired as e:
        sh(["docker", "kill", name])
        out, err = (e.stdout or b"").decode(errors="ignore") if isinstance(e.stdout, bytes) else (e.stdout or ""), ""
        meta.update(timeout=True, exit=None)
    meta["seconds"] = round(time.time() - start, 1)
    (cdir / "agent_stdout.txt").write_text(out or "")
    (cdir / "agent_stderr.txt").write_text(err or "")
    (cdir / "meta.json").write_text(json.dumps(meta))
    return case["id"], meta


def grade(run_dir, no_judge=False):
    cmd = ["docker", "run", "--rm", "--network", "docgen", "--env-file", str(ROOT / ".env"),
           "-e", "GOTENBERG_URL=http://gotenberg:3000", "--entrypoint", "python",
           "-v", f"{EVALS}:/evals:ro", "-v", f"{run_dir}:/run", IMAGE,
           "/evals/evaluate.py", "grade", "/run", "--cases", "/evals/cases.csv"] + (["--no-judge"] if no_judge else [])
    return subprocess.run(cmd).returncode


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cases", default=str(EVALS / "cases.csv"))
    ap.add_argument("--filter", help="regex on case id, e.g. 'PPTX|XLSX-structure'")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--concurrency", type=int, default=4)
    ap.add_argument("--timeout", type=int, default=900, help="seconds per case")
    ap.add_argument("--run-dir")
    ap.add_argument("--grade-only")
    ap.add_argument("--no-judge", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--resume", action="store_true", help="skip cases of --run-dir that already finished")
    a = ap.parse_args()

    if a.selftest:
        ensure_gotenberg()
        return subprocess.run(["docker", "run", "--rm", "--network", "docgen", "-e", "GOTENBERG_URL=http://gotenberg:3000",
                               "--entrypoint", "python", "-v", f"{EVALS}:/evals:ro", IMAGE, "/evals/evaluate.py",
                               "selftest", "--cases", "/evals/cases.csv"] + (["--filter", a.filter] if a.filter else [])).returncode
    if a.grade_only:
        return grade(Path(a.grade_only).resolve(), a.no_judge)
    if not (ROOT / ".env").exists():
        print(".env missing (MODEL_* for the agent, GEMMA_* for the judge)", file=sys.stderr)
        return 2
    run_dir = Path(a.run_dir or EVALS / "runs" / datetime.now().strftime("%Y%m%d-%H%M%S")).resolve()
    run_dir.mkdir(parents=True, exist_ok=True)
    if not ensure_gotenberg():
        print("Gotenberg is not reachable on the docgen network", file=sys.stderr)
        return 1
    cases = load_cases(a.cases, a.filter, a.limit)
    if a.resume:
        def finished(c):
            m = run_dir / c["id"] / "meta.json"
            return m.exists() and not json.loads(m.read_text()).get("timeout")
        skipped = [c["id"] for c in cases if finished(c)]
        cases = [c for c in cases if c["id"] not in skipped]
        print(f"resume: {len(skipped)} finished case(s) kept", flush=True)
    print(f"{len(cases)} cases -> {run_dir}  (concurrency {a.concurrency})", flush=True)
    done = 0
    with cf.ThreadPoolExecutor(max_workers=a.concurrency) as ex:
        futs = [ex.submit(run_case, c, run_dir, a.timeout) for c in cases]
        for fu in cf.as_completed(futs):
            cid, meta = fu.result()
            done += 1
            print(f"[{done}/{len(cases)}] {cid}  {meta['seconds']}s{'  TIMEOUT' if meta['timeout'] else ''}", flush=True)
    print("grading ...", flush=True)
    grade(run_dir, a.no_judge)
    print(f"report: {run_dir / 'report.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
