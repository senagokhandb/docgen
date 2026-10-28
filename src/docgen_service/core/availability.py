"""Availability of the service: the startup self-check and the status behind it.

A job needs the docker daemon, the agent image, Gotenberg on the agent network, the model
endpoint and the agent's model configuration. Which of them are required comes from
configuration (``required_dependencies``). :class:`AvailabilityService` checks them at
start and again when its last result is older than ``availability_ttl_seconds``; the
health endpoint, the MCP status tool and resource and the tools themselves read the same
result, so the host system can switch the capability off instead of planning steps.

Every check has a timeout and never raises. A reason is a short code the host can act
on, with a detail for the operator; neither holds a secret or document content.
"""

from __future__ import annotations

import asyncio
import time
from collections.abc import Awaitable, Callable, Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import Literal

import httpx
from pydantic import BaseModel, Field

from docgen_service.config import Settings
from docgen_service.logging_setup import get_logger

log = get_logger(__name__)

Dependency = Literal["docker", "image", "gotenberg", "model", "agent_env"]
DEPENDENCIES: tuple[Dependency, ...] = ("docker", "image", "gotenberg", "model", "agent_env")

REASONS: dict[str, str] = {
    "docker_unavailable": "The docker daemon does not answer.",
    "image_missing": "The agent image is not built or not loaded.",
    "gotenberg_unavailable": "Gotenberg (the render service) does not answer on the agent network.",
    "model_unavailable": "The model endpoint does not answer or is not configured.",
    "agent_env_missing": "The agent's model configuration is missing or incomplete.",
    "starting": "The self-check has not finished yet.",
    "shutting_down": "The service is shutting down.",
}
REASON_OF: dict[str, str] = {
    "docker": "docker_unavailable",
    "image": "image_missing",
    "gotenberg": "gotenberg_unavailable",
    "model": "model_unavailable",
    "agent_env": "agent_env_missing",
}
AGENT_ENV_KEYS = ("MODEL_ENDPOINT", "MODEL_NAME")


class Component(BaseModel):
    """One dependency as the last check found it."""

    name: str
    required: bool
    ok: bool
    detail: str = ""
    latency_ms: float = 0.0


class AvailabilityStatus(BaseModel):
    """What the health endpoint and the MCP status report."""

    status: Literal["available", "unavailable"]
    reasons: list[str] = Field(default_factory=list, description="Reason codes; empty when available.")
    components: list[Component] = Field(default_factory=list)
    checked_at: datetime | None = None
    version: str = ""

    @property
    def available(self) -> bool:
        return self.status == "available"


CheckFn = Callable[[Settings], Awaitable[tuple[bool, str]]]


def read_env_file(path: Path) -> dict[str, str]:
    """KEY=VALUE lines of an env file (comments and blanks skipped; quotes removed)."""
    values: dict[str, str] = {}
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return values
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _sep, value = line.partition("=")
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key.strip()] = value
    return values


async def run_command(args: list[str], seconds: float) -> tuple[int | None, str]:
    """(exit code, stdout+stderr) of a command; None when it could not run in time."""
    try:
        proc = await asyncio.create_subprocess_exec(
            *args, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.STDOUT
        )
    except OSError as exc:
        return None, f"cannot run {args[0]} ({exc.__class__.__name__})"
    try:
        async with asyncio.timeout(seconds):
            out, _ = await proc.communicate()
    except TimeoutError:
        proc.kill()
        await proc.wait()
        return None, f"no answer within {seconds:g} s"
    return proc.returncode, out.decode("utf-8", "replace").strip()


async def check_docker(settings: Settings) -> tuple[bool, str]:
    code, out = await run_command(
        [settings.docker_bin, "version", "--format", "{{.Server.Version}}"], settings.check_timeout_seconds
    )
    return (True, f"server {out}") if code == 0 else (False, out[-200:] or "no answer")


async def check_image(settings: Settings) -> tuple[bool, str]:
    code, out = await run_command(
        [settings.docker_bin, "image", "inspect", "--format", "{{.Created}}", settings.agent_image],
        settings.check_timeout_seconds,
    )
    return (True, f"{settings.agent_image}, built {out[:19]}") if code == 0 else (False, f"{settings.agent_image} not found")


async def check_gotenberg(settings: Settings) -> tuple[bool, str]:
    url = f"{settings.gotenberg_url.rstrip('/')}/health"
    try:  # the service shares the agent network (container deployment)
        async with httpx.AsyncClient(timeout=3) as client:
            if (await client.get(url)).status_code == 200:
                return True, "answered"
    except httpx.HTTPError:
        pass
    # On the docker host the name only resolves on the agent network: ask from there.
    code, out = await run_command(
        [settings.docker_bin, "run", "--rm", "--network", settings.docker_network, "--entrypoint", "curl",
         settings.agent_image, "-sf", "--max-time", "5", url],
        settings.check_timeout_seconds + 10,
    )
    return (True, "answered on the agent network") if code == 0 else (False, "no answer on the agent network")


async def check_agent_env(settings: Settings) -> tuple[bool, str]:
    values = read_env_file(settings.agent_env_file)
    if not values:
        return False, "file missing or empty"
    missing = [key for key in AGENT_ENV_KEYS if not values.get(key)]
    return (False, "missing " + ", ".join(missing)) if missing else (True, "complete")


async def check_model(settings: Settings) -> tuple[bool, str]:
    values = read_env_file(settings.agent_env_file)
    endpoint, name = values.get("MODEL_ENDPOINT", "").rstrip("/"), values.get("MODEL_NAME", "")
    if not endpoint:
        return False, "MODEL_ENDPOINT is not configured"
    headers = {"Authorization": f"Bearer {values['MODEL_API_KEY']}"} if values.get("MODEL_API_KEY") else {}
    try:
        async with httpx.AsyncClient(timeout=settings.check_timeout_seconds) as client:
            response = await client.get(f"{endpoint}/models", headers=headers)
    except httpx.TimeoutException:
        return False, f"no answer within {settings.check_timeout_seconds:g} s"
    except httpx.HTTPError as exc:
        return False, f"unreachable ({exc.__class__.__name__})"
    if response.status_code != 200:
        return False, f"HTTP {response.status_code}"
    try:
        served = [m.get("id") for m in response.json().get("data", [])]
    except ValueError:
        served = []
    if name and served and name not in served:
        return False, "the configured model is not served"
    return True, "answered"


CHECKS: dict[str, CheckFn] = {
    "docker": check_docker,
    "image": check_image,
    "gotenberg": check_gotenberg,
    "model": check_model,
    "agent_env": check_agent_env,
}


class AvailabilityService:
    """Runs the self-check and keeps its last result for ``availability_ttl_seconds``."""

    def __init__(
        self,
        settings: Settings,
        *,
        version: str = "",
        checks: Mapping[str, CheckFn] | None = None,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._settings = settings
        self._version = version
        self._checks = dict(checks or CHECKS)
        self._clock = clock
        self._last: AvailabilityStatus | None = None
        self._last_at = 0.0
        self._stopping = False
        self._lock: asyncio.Lock | None = None

    def shutting_down(self) -> None:
        self._stopping = True

    def peek(self) -> AvailabilityStatus:
        """The last result without a check (``starting`` before the first one)."""
        if self._stopping:
            return self._with_reason("shutting_down")
        return self._last if self._last is not None else self._with_reason("starting")

    def _with_reason(self, reason: str) -> AvailabilityStatus:
        last = self._last
        return AvailabilityStatus(
            status="unavailable",
            reasons=[reason],
            components=list(last.components) if last else [],
            checked_at=last.checked_at if last else None,
            version=self._version,
        )

    async def status(self, *, refresh: bool = False) -> AvailabilityStatus:
        """The current status; checks again when the last result is too old."""
        if self._stopping:
            return self._with_reason("shutting_down")
        if self._lock is None:  # created inside the running event loop
            self._lock = asyncio.Lock()
        async with self._lock:
            fresh = self._last is not None and self._clock() - self._last_at < self._settings.availability_ttl_seconds
            if refresh or not fresh:
                self._last = await self.check()
                self._last_at = self._clock()
        return self._last if not self._stopping else self._with_reason("shutting_down")

    async def check(self) -> AvailabilityStatus:
        required = set(self._settings.required_dependencies)

        async def one(name: str) -> Component:
            started = time.perf_counter()
            try:
                ok, detail = await self._checks[name](self._settings)
            except Exception as exc:  # a check never raises into the caller
                ok, detail = False, f"check failed ({exc.__class__.__name__})"
            return Component(
                name=name,
                required=name in required,
                ok=ok,
                detail=detail,
                latency_ms=round((time.perf_counter() - started) * 1000.0, 1),
            )

        components = list(await asyncio.gather(*(one(name) for name in DEPENDENCIES if name in self._checks)))
        reasons = [REASON_OF[c.name] for c in components if c.required and not c.ok]
        status = AvailabilityStatus(
            status="unavailable" if reasons else "available",
            reasons=reasons,
            components=components,
            checked_at=datetime.now(UTC),
            version=self._version,
        )
        log.info("availability.checked", status=status.status, reasons=reasons)
        return status
