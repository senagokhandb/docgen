"""Log lines: JSON objects only (library lines included), no request text by default."""

from __future__ import annotations

import io
import json
import logging

from docgen_service.config import Settings
from docgen_service.logging_setup import configure_logging, get_logger, request_field


def test_every_line_is_json_and_requests_are_lengths() -> None:
    out = io.StringIO()
    settings = Settings(_env_file=None, log_format="json")
    configure_logging(settings, stream=out)
    logging.getLogger("httpx").info("HTTP Request: GET http://gotenberg:3000/health")
    get_logger("test").info("jobs.submitted", job="j1", **request_field(settings, "Gizli rapor"))
    lines = [json.loads(line) for line in out.getvalue().splitlines()]
    assert [line["event"].split(":")[0] for line in lines] == ["HTTP Request", "jobs.submitted"]
    assert lines[1]["request_chars"] == len("Gizli rapor") and "Gizli" not in out.getvalue()
    assert request_field(Settings(_env_file=None, log_request_text=True), "x") == {"request": "x"}
