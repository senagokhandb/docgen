"""Handover: the capability definition does not drift from the MCP server and the code,
the routing test set is well-formed, and the attribution is where it belongs."""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any

import yaml

from docgen_service import __version__
from docgen_service.config import MB, Settings
from docgen_service.conversation import DELIVERABLE, UPLOADABLE
from docgen_service.core.availability import REASONS
from docgen_service.interfaces.mcp_server import (
    BUSY_REASONS,
    ERROR_BUSY,
    ERROR_JOB_FAILED,
    ERROR_NOT_FOUND,
    ERROR_UNAVAILABLE,
    JOB_STATUSES,
    NOT_FOUND_REASONS,
    READ_ERROR_KINDS,
    SERVER_NAME,
    SERVER_TITLE,
    STATUS_URI,
)
from docgen_service.jobs import FAILURE_REASONS, READABLE

from .test_mcp import FakeBackends, _server

ROOT = Path(__file__).parents[2]
CAPABILITY = ROOT / "docs" / "capability" / "document_capability.yaml"
ROUTING = ROOT / "docs" / "capability" / "routing_test_set.yaml"
PROVIDER = "Digital Brain Technologies"
ATTRIBUTION = f"Developed by {PROVIDER}"
JSON_TYPES = {"string": "string", "integer": "integer", "boolean": "boolean", "array": "array"}


def capability() -> dict[str, Any]:
    return yaml.safe_load(CAPABILITY.read_text(encoding="utf-8"))["capability"]


def plain(schema: dict[str, Any]) -> dict[str, Any]:
    """The schema of a parameter without the wrapping of an optional value."""
    options = [o for o in schema.get("anyOf", []) if o.get("type") != "null"]
    return {**schema, **options[0]} if len(options) == 1 else schema


async def test_tools_and_parameters_match_the_mcp_schemas(tmp_path: Path) -> None:
    declared = capability()["tools"]
    served = {tool.name: tool for tool in await _server(FakeBackends(tmp_path)).list_tools()}
    assert set(declared) == set(served)
    for name, tool in served.items():
        wanted = declared[name]["input"] or {}
        properties = tool.input_schema.get("properties", {})
        required = set(tool.input_schema.get("required", []))
        assert set(wanted) == set(properties), name
        for parameter, spec in wanted.items():
            schema = plain(properties[parameter])
            where = f"{name}.{parameter}"
            assert schema.get("type") == JSON_TYPES[spec["type"]], where
            assert bool(spec.get("required", False)) == (parameter in required), where
            if "default" in spec:
                assert properties[parameter].get("default") == spec["default"], where
        outputs = set((tool.output_schema or {}).get("properties", {}))
        named = set(declared[name]["output"])
        assert named <= outputs, (name, sorted(named - outputs))


def test_every_rule_is_plain_text() -> None:
    """A colon in a plain YAML list item turns the rule into a mapping without an error."""

    def walk(node: Any, where: str) -> None:
        if isinstance(node, dict):
            for key, value in node.items():
                walk(value, f"{where}.{key}")
        elif isinstance(node, list):
            for i, item in enumerate(node):
                assert isinstance(item, str | dict) and (isinstance(item, str) or where.endswith("changelog")), (
                    f"{where}[{i}]")
                walk(item, f"{where}[{i}]")

    walk(capability(), "capability")


def test_errors_availability_jobs_and_server_match_the_code() -> None:
    data = capability()
    assert data["version"] == "1.0" and data["provider"] == PROVIDER
    assert data["status"] == "available" and data["transport"] == "mcp"
    assert data["server"] == {"name": SERVER_NAME, "title": SERVER_TITLE}
    assert [entry["version"] for entry in data["changelog"]] == ["1.0"]
    assert all(entry["changes"] for entry in data["changelog"])

    availability = data["availability"]
    assert availability["status_tool"] == "document_status"
    assert availability["status_resource"] == STATUS_URI
    assert availability["statuses"] == ["available", "unavailable"]
    assert set(availability["reasons"]) == set(REASONS)
    assert "/health" in availability["health_endpoint"]

    errors = data["errors"]
    assert {ERROR_UNAVAILABLE, ERROR_BUSY, ERROR_NOT_FOUND, ERROR_JOB_FAILED} <= set(errors)
    assert set(errors[ERROR_UNAVAILABLE]["reasons"]) == set(REASONS)
    assert errors[ERROR_BUSY]["reasons"] == list(BUSY_REASONS)
    assert errors[ERROR_NOT_FOUND]["reasons"] == list(NOT_FOUND_REASONS)
    assert errors[ERROR_JOB_FAILED]["reasons"] == list(FAILURE_REASONS)
    assert errors["read_error"]["kinds"] == list(READ_ERROR_KINDS)
    for name, error in errors.items():
        assert error["meaning"].strip() and error["host_action"].strip(), name

    assert list(data["jobs"]["statuses"]) == list(JOB_STATUSES)
    formats = data["formats"]
    assert set(formats["inputs"]) == UPLOADABLE
    assert set(formats["outputs"]) == DELIVERABLE
    assert set(formats["readable"]) == READABLE
    shipped = Settings(_env_file=None)
    limits = data["files"]["limits"]
    assert limits["files_per_job"] == shipped.max_input_files
    assert limits["bytes_per_file"] == f"{shipped.max_input_file_bytes // MB} MB"
    assert limits["bytes_per_job"] == f"{shipped.max_input_total_bytes // MB} MB"
    assert set(data["performance"]["recommended_timeouts"]) == set(data["tools"])
    assert "only what was asked" in data["guardrails"]["document_agent"]
    assert "host system's" in data["guardrails"]["host"]
    assert "never on the planner's own initiative" in " ".join(data["planning_rules"])
    assert data["dates"]["format"].startswith("RFC 3339 date-time in UTC")


def test_the_routing_test_set_is_complete_and_generic() -> None:
    data = yaml.safe_load(ROUTING.read_text(encoding="utf-8"))
    cases = data["cases"]
    assert data["capability_version"] == capability()["version"]
    assert 38 <= len(cases) <= 50
    assert len({case["id"] for case in cases}) == len(cases)
    kinds = {"document_create", "document_edit", "document_continue", "document_read", "no_document",
             "gather_first", "confirm"}
    routes = {"internal_documents", "business_data", "web_research", "email_calendar", "conversation", "image"}
    by_kind: dict[str, list[dict[str, Any]]] = {kind: [] for kind in kinds}
    for case in cases:
        assert case["language"] in ("tr", "en") and case["request"].strip(), case["id"]
        by_kind[case["expected"]].append(case)
        for name in case.get("files", []):
            assert Path(name).suffix in UPLOADABLE, case["id"]
    for kind, found in by_kind.items():
        assert {c["language"] for c in found} == {"tr", "en"}, kind
        assert len(found) >= 4, kind
    for case in by_kind["document_edit"]:
        assert case["files"], case["id"]
    for case in by_kind["document_continue"]:
        assert "job" in case["context"], case["id"]
    for case in by_kind["no_document"]:
        assert case["route"] in routes, case["id"]
    assert {c["route"] for c in by_kind["no_document"]} == routes
    for case in by_kind["gather_first"]:
        assert case["first_route"] in {"internal_documents", "business_data", "web_research"}, case["id"]
    assert {c["confirm"] for c in by_kind["confirm"]} == {True, False}
    for case in cases:  # names of companies, projects and persons are placeholders
        assert "http" not in case["request"], case["id"]


def test_attribution_is_in_metadata_and_docs() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    assert {"name": PROVIDER} in project["authors"]
    assert ATTRIBUTION in project["description"]
    assert project["version"] == __version__
    assert ATTRIBUTION in (ROOT / "docs" / "INTEGRATION.md").read_text(encoding="utf-8")
    assert SERVER_TITLE == f"Document Agent by {PROVIDER}"
