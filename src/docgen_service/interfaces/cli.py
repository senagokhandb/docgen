"""Command line: run the MCP server, check the status, run one job locally.

    docgen-service mcp --transport http          # MCP over streamable HTTP (deployment)
    docgen-service mcp                           # MCP over stdio (a local MCP client)
    docgen-service status                        # the self-check, as JSON
    docgen-service run "Bir sayfalık bir rapor yap" --file veri.csv --out ./sonuc
"""

from __future__ import annotations

import asyncio
import json
import shutil
from pathlib import Path

import typer

from docgen_service import __version__
from docgen_service.config import get_settings
from docgen_service.core.availability import AvailabilityService
from docgen_service.logging_setup import configure_logging

app = typer.Typer(help="DOCGEN service: document jobs over MCP.", no_args_is_help=True)


@app.command()
def mcp(
    transport: str = typer.Option("stdio", "--transport", help="stdio or http."),
    host: str | None = typer.Option(None, "--host", help="HTTP bind host (default from config)."),
    port: int | None = typer.Option(None, "--port", help="HTTP port (default from config)."),
) -> None:
    """Run the MCP server exposing document_run, document_job, document_cancel, document_read
    and document_status."""
    from docgen_service.interfaces.mcp_server import run_server

    if transport not in ("stdio", "http"):
        typer.echo("error: --transport must be stdio or http", err=True)
        raise typer.Exit(code=2)
    try:
        run_server(transport, host=host, port=port)
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(code=1) from exc


@app.command()
def status() -> None:
    """Run the self-check once and print it; exit code 1 when unavailable."""
    settings = get_settings()
    configure_logging(settings)
    result = asyncio.run(AvailabilityService(settings, version=__version__).check())
    typer.echo(result.model_dump_json(indent=2))
    raise typer.Exit(code=0 if result.available else 1)


@app.command()
def run(
    request: str = typer.Argument(..., help="What the agent should do."),
    file: list[Path] = typer.Option([], "--file", "-f", help="An input file (repeatable)."),
    confirm: bool = typer.Option(False, "--confirm", help="The user agreed to a destructive change."),
    out: Path = typer.Option(Path("."), "--out", help="Where the delivered files are copied."),
) -> None:
    """Run one job here and wait for it (no MCP): for operators and smoke tests."""
    from docgen_service.jobs import JobService

    settings = get_settings()
    configure_logging(settings)

    async def go() -> int:
        jobs = JobService(settings)
        await jobs.start()
        try:
            job = jobs.submit(request, [(p.name, p.read_bytes()) for p in file], confirm=confirm)
            await jobs.wait(job, float(settings.job_timeout_seconds + 60))
        finally:
            await jobs.close()
        out.mkdir(parents=True, exist_ok=True)
        for record in job.outputs:
            shutil.copyfile(jobs.output_path(job, record.name), out / record.name)
        typer.echo(json.dumps({"job_id": job.id, "status": job.status, "files": [f.name for f in job.outputs],
                               "reply": job.reply, "failure": job.failure, "message": job.message},
                              ensure_ascii=False, indent=2))
        return 0 if job.status == "done" else 1

    raise typer.Exit(code=asyncio.run(go()))


def main() -> None:
    app()
