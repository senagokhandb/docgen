"""Configuration of the DOCGEN service.

Settings are read from environment variables (prefix ``DOCGEN_``) and, when
``DOCGEN_ENV_FILE`` names one, from that file. They are kept apart from the agent's
``.env``: every value of that file is passed into the agent containers, so the
service's own secrets (the MCP token) must never be written there.
"""

from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Annotated, Literal

from pydantic import Field, SecretStr, field_validator
from pydantic_settings import BaseSettings, NoDecode, SettingsConfigDict

LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
LogFormat = Literal["console", "json"]

# List settings are given as comma-separated strings in the environment.
StrList = Annotated[list[str], NoDecode]

MB = 1024 * 1024


class Settings(BaseSettings):
    """All runtime configuration of the service."""

    model_config = SettingsConfigDict(
        env_prefix="DOCGEN_",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )

    # --- Agent containers ---
    agent_image: str = Field(
        default="docgen-agent:latest", description="Image every job runs in (Dockerfile)."
    )
    agent_env_file: Path = Field(
        default=Path(".env"),
        description=(
            "The agent's model configuration (MODEL_ENDPOINT, MODEL_NAME, MODEL_API_KEY), "
            "passed to every job container with --env-file."
        ),
    )
    docker_bin: str = Field(default="docker", description="The docker client.")
    docker_network: str = Field(
        default="docgen", description="Network of the job containers; Gotenberg is on it."
    )
    gotenberg_url: str = Field(
        default="http://gotenberg:3000",
        description="Gotenberg as the job containers reach it (on docker_network).",
    )

    # --- Jobs ---
    jobs_dir: Path = Field(
        default=Path("data/mcp-jobs"),
        description="Where the workspace and the record of every job are kept.",
    )
    jobs_host_dir: Path | None = Field(
        default=None,
        description=(
            "jobs_dir as the docker daemon sees it, for the bind mounts of the job "
            "containers. Empty: jobs_dir (the service runs on the docker host). Set it when "
            "the service runs in a container with the host's docker socket."
        ),
    )
    max_running_jobs: int = Field(
        default=3, ge=1, le=32, description="Agent containers running at the same time."
    )
    max_queued_jobs: int = Field(
        default=20,
        ge=0,
        description="Jobs waiting for a place; a new job beyond this is refused with busy.",
    )
    job_timeout_seconds: int = Field(
        default=1500, ge=60, description="A job still running after this is stopped (timed_out)."
    )
    job_memory_limit: str = Field(
        default="", description="docker --memory of a job container, for example 4g. Empty: none."
    )
    job_cpus: str = Field(
        default="", description="docker --cpus of a job container, for example 2. Empty: none."
    )
    job_retention_hours: float = Field(
        default=24.0,
        gt=0,
        description="Finished jobs (workspace and files) are deleted after this many hours.",
    )
    max_input_files: int = Field(default=10, ge=1)
    max_input_file_bytes: int = Field(default=25 * MB, ge=1)
    max_input_total_bytes: int = Field(default=50 * MB, ge=1)
    inline_output_max_bytes: int = Field(
        default=10 * MB,
        ge=0,
        description="Largest output file returned inline (include_content) as base64.",
    )
    read_timeout_seconds: int = Field(
        default=90, ge=5, description="Time document_read may take (one container run)."
    )
    read_max_chars: int = Field(default=60_000, ge=1000)

    # --- MCP server ---
    mcp_http_host: str = Field(default="127.0.0.1")
    mcp_http_port: int = Field(default=8766, ge=1, le=65535)
    mcp_http_token: SecretStr = Field(
        default=SecretStr(""),
        description="Bearer token required by the HTTP transport (mandatory off localhost).",
    )
    mcp_max_concurrent_calls: int = Field(
        default=16,
        ge=1,
        le=256,
        description="Tool calls of every kind served at the same time; others wait.",
    )
    mcp_queue_timeout_seconds: float = Field(
        default=30.0,
        ge=0,
        description="How long a tool call waits for a free place before it returns busy.",
    )
    mcp_max_wait_seconds: int = Field(
        default=110,
        ge=0,
        le=600,
        description=(
            "Upper bound of wait_seconds of document_run and document_job: a call never "
            "holds the connection longer than this (clients time out around 120 s)."
        ),
    )
    mcp_max_request_bytes: int = Field(
        default=80 * MB,
        ge=MB,
        description="Largest HTTP request body (inputs arrive base64 encoded: about 4/3 larger).",
    )
    mcp_shutdown_grace_seconds: int = Field(
        default=30,
        ge=1,
        description=(
            "Time running tool calls get to return when the service is stopped. Running "
            "jobs are then stopped and recorded as interrupted."
        ),
    )

    # --- Availability ---
    required_dependencies: StrList = Field(
        default_factory=lambda: ["docker", "image", "gotenberg", "model", "agent_env"],
        description=(
            "Dependencies without which the service reports itself unavailable: docker "
            "(the daemon answers), image (agent_image exists), gotenberg, model (the model "
            "endpoint of agent_env_file answers), agent_env (agent_env_file is complete)."
        ),
    )
    availability_ttl_seconds: float = Field(
        default=30.0,
        ge=0,
        description="How long the result of the self-check is reused before it runs again.",
    )
    startup_check_required: bool = Field(
        default=False,
        description=(
            "Refuse to start when the startup self-check finds a required dependency "
            "missing. Off: the server starts and reports itself unavailable."
        ),
    )
    check_timeout_seconds: float = Field(default=10.0, gt=0)

    # --- Privacy and logging ---
    log_request_text: bool = Field(
        default=False,
        description=(
            "Write the text of requests into our own logs. Off: their length stands in "
            "their place. Document content is never logged."
        ),
    )
    log_level: LogLevel = Field(default="INFO", description="Minimum log level.")
    log_format: LogFormat = Field(
        default="console", description="'console' for humans, 'json' for log pipelines."
    )
    log_stream: Literal["stderr", "stdout"] = Field(
        default="stderr",
        description=(
            "Where the log lines go. The MCP stdio transport needs stdout for itself and "
            "always logs to stderr."
        ),
    )

    @field_validator("required_dependencies", mode="before")
    @classmethod
    def _split_lists(cls, value: object) -> object:
        """Comma-separated in the environment: 'docker,image,model'."""
        if isinstance(value, str):
            return [part.strip() for part in value.split(",") if part.strip()]
        return value

    @property
    def jobs_mount_dir(self) -> Path:
        """jobs_dir as the docker daemon sees it."""
        return self.jobs_host_dir or self.jobs_dir.resolve()


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    env_file = os.environ.get("DOCGEN_ENV_FILE") or None
    return Settings(_env_file=env_file)
