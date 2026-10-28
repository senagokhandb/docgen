"""MCP contract: every tool is called with fixtures and its structured output is validated
against the output schema the tool declares, formats included. One date representation
across the tools: RFC 3339 date-time in UTC with the suffix Z."""

from __future__ import annotations

import base64
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import jsonschema
from mcp.client.session import ClientSession

from docgen_service.interfaces.mcp_server import rfc3339

from .test_mcp import FakeBackends, _server, in_process_session

DATE_CHARS = len("2026-10-28T09:30:00Z")
CHECKER = jsonschema.FormatChecker()


@CHECKER.checks("date-time")
def is_rfc3339_utc(value: object) -> bool:
    """The documented representation, nothing else: YYYY-MM-DDTHH:MM:SSZ."""
    if not isinstance(value, str):
        return True
    if len(value) != DATE_CHARS or not value.endswith("Z") or value[10] != "T":
        return False
    try:
        datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return False
    return True


def validate(output: Any, schema: dict[str, Any]) -> None:
    validator = jsonschema.Draft202012Validator(schema, format_checker=CHECKER)
    errors = sorted(validator.iter_errors(output), key=lambda e: list(e.path))
    assert errors == [], [f"{list(e.path)}: {e.message}" for e in errors]


def date_formats(schema: Any) -> int:
    """How many fields of ``schema`` are declared as date-time."""
    if isinstance(schema, dict):
        own = 1 if schema.get("format") == "date-time" else 0
        return own + sum(date_formats(v) for v in schema.values())
    if isinstance(schema, list):
        return sum(date_formats(v) for v in schema)
    return 0


async def checked(session: ClientSession, schemas: dict[str, Any], name: str, args: dict) -> dict:
    result = await session.call_tool(name, args)
    assert not result.is_error, result.content
    validate(result.structured_content, schemas[name])
    return result.structured_content


def test_rfc3339_is_utc_seconds_z() -> None:
    istanbul = datetime(2026, 10, 28, 12, 30, 15, 250000, tzinfo=timezone(timedelta(hours=3)))
    assert rfc3339(istanbul) == "2026-10-28T09:30:15Z"
    assert rfc3339(datetime(2026, 10, 28)) == "2026-10-28T00:00:00Z"


async def test_every_output_matches_its_declared_schema(tmp_path: Path) -> None:
    fake = FakeBackends(tmp_path / "up")
    async with in_process_session(_server(fake)) as session:
        tools = {t.name: t for t in (await session.list_tools()).tools}
        schemas = {name: tool.output_schema for name, tool in tools.items()}
        assert date_formats(schemas["document_run"]) == 3
        assert date_formats(schemas["document_status"]) == 1
        docx = base64.b64encode(b"PK fake").decode()
        done = await checked(session, schemas, "document_run",
                             {"request": "Rapor yap", "files": [{"name": "a.docx", "content_base64": docx}],
                              "include_content": True})
        assert done["status"] == "done" and done["created_at"].endswith("Z")
        await checked(session, schemas, "document_run", {"request": "FAIL"})
        await checked(session, schemas, "document_run", {"request": "REFUSE"})
        running = await checked(session, schemas, "document_run", {"request": "HANG", "wait_seconds": 0})
        await checked(session, schemas, "document_job", {"job_id": running["job_id"], "wait_seconds": 0})
        fake.agent.release.set()
        await checked(session, schemas, "document_job", {"job_id": running["job_id"], "wait_seconds": 5})
        await checked(session, schemas, "document_job", {"job_id": "unknown"})
        await checked(session, schemas, "document_cancel", {"job_id": done["job_id"]})
        await checked(session, schemas, "document_cancel", {"job_id": "unknown"})
        await checked(session, schemas, "document_read", {"name": "a.pdf", "content_base64": docx})
        await checked(session, schemas, "document_read", {"name": "a.exe", "content_base64": docx})
        status = await checked(session, schemas, "document_status", {})
        assert status["checked_at"].endswith("Z")
    down = FakeBackends(tmp_path / "down", down=("docker",))
    async with in_process_session(_server(down)) as session:
        schemas = {t.name: t.output_schema for t in (await session.list_tools()).tools}
        refused = await checked(session, schemas, "document_run", {"request": "Rapor yap"})
        assert refused["status"] == "not_started" and refused["created_at"] is None
        await checked(session, schemas, "document_read", {"name": "a.pdf", "content_base64": docx})
        await checked(session, schemas, "document_status", {})
