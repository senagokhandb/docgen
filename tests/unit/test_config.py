"""Settings as a deployment gives them: environment variables and an env file."""

from __future__ import annotations

from pathlib import Path

import pytest

from docgen_service.config import Settings, get_settings
from docgen_service.core.availability import read_env_file


def test_settings_from_the_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCGEN_REQUIRED_DEPENDENCIES", "docker, image,model")
    monkeypatch.setenv("DOCGEN_MAX_RUNNING_JOBS", "2")
    monkeypatch.setenv("DOCGEN_MCP_HTTP_TOKEN", "t0ken")
    monkeypatch.setenv("DOCGEN_JOBS_HOST_DIR", "/srv/docgen/jobs")
    settings = Settings(_env_file=None)
    assert settings.required_dependencies == ["docker", "image", "model"]
    assert settings.max_running_jobs == 2
    assert settings.mcp_http_token.get_secret_value() == "t0ken"
    assert "t0ken" not in repr(settings)
    assert settings.jobs_mount_dir == Path("/srv/docgen/jobs")


def test_the_env_file_named_by_docgen_env_file(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    env = tmp_path / "service.env"
    env.write_text("DOCGEN_REQUIRED_DEPENDENCIES=docker,image\nDOCGEN_JOB_TIMEOUT_SECONDS=600\n")
    monkeypatch.setenv("DOCGEN_ENV_FILE", str(env))
    get_settings.cache_clear()
    try:
        settings = get_settings()
    finally:
        get_settings.cache_clear()
    assert settings.required_dependencies == ["docker", "image"] and settings.job_timeout_seconds == 600


def test_the_shipped_examples_load(monkeypatch: pytest.MonkeyPatch) -> None:
    root = Path(__file__).parents[2]
    for key, value in read_env_file(root / "deploy" / "service.env.example").items():
        monkeypatch.setenv(key, value)
    settings = Settings(_env_file=None)
    assert settings.required_dependencies == ["docker", "image", "gotenberg", "model", "agent_env"]
    assert settings.job_memory_limit == "" and settings.max_running_jobs == 3


def test_read_env_file_handles_comments_and_quotes(tmp_path: Path) -> None:
    env = tmp_path / ".env"
    env.write_text('# model\nMODEL_ENDPOINT="http://m:8000/v1"\n\nMODEL_NAME=q\nBROKEN\n')
    assert read_env_file(env) == {"MODEL_ENDPOINT": "http://m:8000/v1", "MODEL_NAME": "q"}
    assert read_env_file(tmp_path / "missing") == {}
