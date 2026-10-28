"""Structured logging (structlog): console lines for humans, JSON lines for pipelines.

Log events carry identifiers, sizes, timings and reason codes. They never carry the
content of documents; request text only with ``log_request_text``.
"""

from __future__ import annotations

import logging
import sys
from typing import Any, TextIO

import structlog

from docgen_service.config import Settings


def configure_logging(settings: Settings, stream: TextIO | None = None) -> None:
    out = stream or (sys.stdout if settings.log_stream == "stdout" else sys.stderr)
    level = getattr(logging, settings.log_level)
    renderer: Any = (
        structlog.processors.JSONRenderer(ensure_ascii=False)
        if settings.log_format == "json"
        else structlog.dev.ConsoleRenderer(colors=False)
    )
    # Lines of libraries (httpx, the MCP SDK, uvicorn) get the same format: one JSON object
    # per line in a container, never a plain line in between.
    handler = logging.StreamHandler(out)
    handler.setFormatter(structlog.stdlib.ProcessorFormatter(
        processor=renderer,
        foreign_pre_chain=[structlog.stdlib.add_log_level, structlog.stdlib.add_logger_name,
                           structlog.processors.TimeStamper(fmt="iso", utc=True)],
    ))
    logging.basicConfig(handlers=[handler], level=level, force=True)
    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            structlog.processors.add_log_level,
            structlog.processors.TimeStamper(fmt="iso", utc=True),
            renderer,
        ],
        wrapper_class=structlog.make_filtering_bound_logger(level),
        logger_factory=structlog.PrintLoggerFactory(file=out),
        cache_logger_on_first_use=False,
    )


def get_logger(name: str) -> Any:
    return structlog.get_logger(name)


def request_field(settings: Settings, text: str) -> dict[str, Any]:
    """How a request appears in a log line: its text only when configured."""
    return {"request": text} if settings.log_request_text else {"request_chars": len(text)}
