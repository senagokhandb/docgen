"""Automated MCP client against a running document service over HTTP.

Lists the tools, calls every tool, validates every structured output against the output
schema the tool declares (formats included), runs a real job end to end (create a Word
document, download it as a resource and over HTTP, edit it with continue_from, read it
back) unless --no-jobs. Exit code 0 when everything passed.

    python deploy/check_mcp.py --url http://127.0.0.1:18766 --token-file deploy/stack.env
    python deploy/check_mcp.py --no-jobs            # protocol checks only (no model time)

Needs the packages mcp, httpx and jsonschema (pip install -e ".[dev]").
"""

from __future__ import annotations

import argparse
import asyncio
import base64
import hashlib
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
import jsonschema
from mcp.client.session import ClientSession
from mcp.client.streamable_http import create_mcp_http_client, streamable_http_client

TOKEN_KEY = "DOCGEN_MCP_HTTP_TOKEN"
DATE_CHARS = len("2026-10-28T09:30:00Z")
CHECKER = jsonschema.FormatChecker()
WANTED = {"document_run", "document_job", "document_cancel", "document_read", "document_status"}
CREATE = "Başlığı 'MCP Testi' olan, iki kısa paragraflık tek sayfalık bir Word belgesi oluştur."
EDIT = "Başlığı kırmızı yap."


@CHECKER.checks("date-time")
def is_rfc3339_utc(value: object) -> bool:
    if not isinstance(value, str):
        return True
    if len(value) != DATE_CHARS or not value.endswith("Z") or value[10] != "T":
        return False
    try:
        datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return False
    return True


def read_token(path: Path) -> str:
    for line in path.read_text(encoding="utf-8").splitlines():
        name, _equals, value = line.partition("=")
        if name.strip() == TOKEN_KEY:
            return value.strip()
    raise SystemExit(f"{TOKEN_KEY} not found in {path}")


def problems(output: Any, schema: dict[str, Any]) -> list[str]:
    validator = jsonschema.Draft202012Validator(schema, format_checker=CHECKER)
    return [f"{list(e.path)}: {e.message[:120]}" for e in validator.iter_errors(output)]


class Report:
    def __init__(self) -> None:
        self.rows: list[dict[str, Any]] = []

    def add(self, check: str, ok: bool, detail: str = "", seconds: float = 0.0) -> None:
        self.rows.append({"check": check, "ok": ok, "detail": detail, "seconds": round(seconds, 1)})
        print(f"{'ok  ' if ok else 'FAIL'}  {check}  {detail}  ({seconds:.1f} s)", flush=True)

    @property
    def failed(self) -> int:
        return sum(1 for row in self.rows if not row["ok"])


async def run(args: argparse.Namespace, token: str) -> Report:
    report = Report()
    base = args.url.rstrip("/")
    auth = {"Authorization": f"Bearer {token}"}
    async with httpx.AsyncClient(timeout=30) as plain:
        started = time.perf_counter()
        health = await plain.get(f"{base}/health")
        body = health.json()
        report.add("health endpoint", health.status_code == 200,
                   f"HTTP {health.status_code}, {body.get('status')}, reasons {body.get('reasons')}",
                   time.perf_counter() - started)
        refused = await plain.post(f"{base}/mcp", json={})
        report.add("no token is refused", refused.status_code == 401, f"HTTP {refused.status_code}")
        download = await plain.get(f"{base}/jobs/x/files/y")
        report.add("downloads need the token", download.status_code == 401, f"HTTP {download.status_code}")

    http = create_mcp_http_client(headers=auth)
    async with streamable_http_client(f"{base}/mcp", http_client=http) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            info = session.server_info
            report.add("server info", bool(info.name and info.title and info.version),
                       f"{info.name} / {info.title} / {info.version}")
            tools = {t.name: t for t in (await session.list_tools()).tools}
            report.add("tools listed", set(tools) == WANTED, ", ".join(sorted(tools)))
            report.add("every tool declares an output schema", all(t.output_schema for t in tools.values()))

            async def check(label: str, name: str, arguments: dict, test: Any, timeout: float = 130) -> Any:
                print(f"      running {label} ...", flush=True)
                started = time.perf_counter()
                try:
                    async with asyncio.timeout(timeout):
                        result = await session.call_tool(name, arguments)
                    if result.is_error:
                        raise RuntimeError(f"tool error: {result.content[0].text[:200]}")
                    output = result.structured_content
                    found = problems(output, tools[name].output_schema)
                    if found:
                        raise RuntimeError("output does not match the declared schema: " + "; ".join(found[:3]))
                    ok, detail = test(output)
                except Exception as exc:
                    output, ok, detail = None, False, f"{exc.__class__.__name__}: {exc}"
                report.add(label, ok, detail, time.perf_counter() - started)
                return output

            status = await check("document_status", "document_status", {},
                                 lambda o: (o["status"] == "available",
                                            f"{o['status']}, reasons {o['reasons']}, jobs {o['jobs']}"))
            resource = await session.read_resource("docgen://status")
            same = json.loads(resource.contents[0].text)["status"] == (status or {}).get("status")
            report.add("status resource", same, "docgen://status")
            await check("unknown job is not_found", "document_job", {"job_id": "nope", "wait_seconds": 0},
                        lambda o: ((o.get("error") or {}).get("kind") == "not_found", f"error {o.get('error')}"))
            bad = await session.call_tool("document_run", {"request": " "})
            report.add("empty request is a tool error", bad.is_error, bad.content[0].text[:80] if bad.content else "")
            if args.no_jobs:
                return report

            async def finish(label: str, first: dict | None) -> dict | None:
                output = first
                deadline = time.monotonic() + args.job_timeout
                while output and not output["finished"] and time.monotonic() < deadline:
                    output = await check(f"{label}: document_job", "document_job",
                                         {"job_id": output["job_id"], "wait_seconds": 100},
                                         lambda o: (o["job_id"] != "", f"{o['status']}, {o['step_count']} steps, "
                                                    f"{o['current_activity'][:60]}"))
                return output

            def delivered(o: dict) -> tuple[bool, str]:
                names = [f["name"] for f in o["files"]]
                return (o["status"] == "done" and any(n.endswith(".docx") for n in names),
                        f"{o['status']}, files {names}, {o['total_ms'] / 1000:.0f} s, reply {o['reply'][:80]!r}")

            created = await finish("create", await check(
                "document_run (create)", "document_run", {"request": CREATE, "wait_seconds": 100},
                lambda o: (o["job_id"] != "" and o["error"] is None, f"{o['status']} {o['job_id']}")))
            ok, detail = delivered(created) if created else (False, "no job")
            report.add("create: a .docx was delivered", ok, detail)
            if not ok:
                return report
            file = next(f for f in created["files"] if f["name"].endswith(".docx"))
            blob = base64.b64decode((await session.read_resource(file["uri"])).contents[0].blob)
            report.add("file as MCP resource", hashlib.sha256(blob).hexdigest() == file["sha256"],
                       f"{file['uri']} {len(blob)} bytes")
            async with httpx.AsyncClient(timeout=60) as plain:
                got = await plain.get(base + file["download_path"], headers=auth)
            report.add("file over HTTP", got.status_code == 200 and hashlib.sha256(got.content).hexdigest()
                       == file["sha256"], f"HTTP {got.status_code}, {len(got.content)} bytes")

            edited = await finish("edit", await check(
                "document_run (continue_from)", "document_run",
                {"request": EDIT, "continue_from": created["job_id"], "wait_seconds": 100},
                lambda o: (o["chain_length"] == 2 and o["error"] is None, f"{o['status']} {o['job_id']}")))
            ok, detail = delivered(edited) if edited else (False, "no job")
            report.add("edit: a .docx was delivered", ok, detail)
            if ok:
                name = next(f["name"] for f in edited["files"] if f["name"].endswith(".docx"))
                verified = [v for v in edited["verification"] if v["ok"]]
                report.add("edit: diff check recorded", bool(verified), f"{edited['verification'][:1]}")
                await check("document_read (formatting)", "document_read",
                            {"name": name, "job_id": edited["job_id"], "formatting": True},
                            lambda o: (o["ok"] and "MCP Testi" in o["untrusted_text"],
                                       f"{o['chars']} characters; title line: "
                                       + next((ln.strip() for ln in o["untrusted_text"].splitlines()
                                               if "MCP Testi" in ln), "-")[:120]))
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--url", default="http://127.0.0.1:18766")
    parser.add_argument("--token-file", default="deploy/stack.env")
    parser.add_argument("--token-env", action="store_true",
                        help=f"Take the token from the environment variable {TOKEN_KEY} (no file is read)")
    parser.add_argument("--no-jobs", action="store_true", help="Protocol checks only: no job is started")
    parser.add_argument("--job-timeout", type=float, default=900, help="Seconds a job may take")
    parser.add_argument("--out", default="", help="Write the checks as JSON to this file")
    args = parser.parse_args()
    token = os.environ.get(TOKEN_KEY, "") if args.token_env else read_token(Path(args.token_file))
    if not token:
        raise SystemExit(f"{TOKEN_KEY} is empty")
    report = asyncio.run(run(args, token))
    if args.out:
        Path(args.out).write_text(json.dumps(report.rows, indent=1), encoding="utf-8")
    print(f"{len(report.rows) - report.failed} of {len(report.rows)} checks passed")
    sys.exit(1 if report.failed else 0)


if __name__ == "__main__":
    main()
