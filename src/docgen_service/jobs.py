"""Document jobs: one request of the user, run by the agent in a throw-away container.

A job gets its own workspace (``uploads/`` read-only, ``scratch/``, ``outputs/``) under
``jobs_dir`` and runs the docgen-agent image exactly like the chat UI and the eval do:
same entrypoint, same network with Gotenberg, the agent's ``.env``, stream-json events.
A job may continue another one (``continue_from``): the files and the scratch/ of the
whole chain come along, and the prompt is the chat UI's (conversation.build_prompt), so
"make the title red" edits the file the previous job delivered.

Jobs are kept on disk (``job.json``) for ``job_retention_hours``; a job that was running
when the service stopped is recorded as interrupted. Never raises into the MCP layer for
a failed run: the outcome is the status of the job.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import mimetypes
import os
import shutil
import time
import uuid
from collections.abc import Awaitable, Callable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field

from docgen_service import conversation as conv
from docgen_service.config import Settings
from docgen_service.core.availability import run_command
from docgen_service.logging_setup import get_logger, request_field

log = get_logger(__name__)

JobStatus = Literal["queued", "running", "done", "failed", "timed_out", "cancelled", "interrupted"]
FINISHED: frozenset[str] = frozenset({"done", "failed", "timed_out", "cancelled", "interrupted"})
FAILURE_REASONS = ("agent_error", "timed_out", "interrupted")
READABLE = {".docx", ".pptx", ".xlsx", ".pdf", ".doc", ".ppt", ".xls",
            ".dotx", ".docm", ".potx", ".pptm", ".xltx", ".xlsm"}
FORMAT_READABLE = {".docx", ".pptx", ".xlsx", ".pdf"}
CONTAINER_PREFIX = "docgen-job-"
JOB_LABEL = "docgen.job"
MANIFEST = "manifest.json"
PROGRESS_SECONDS = 2.0

for _ext, _type in ((".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                    (".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
                    (".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    (".md", "text/markdown")):
    mimetypes.add_type(_type, _ext)


class InvalidInput(ValueError):
    """The call itself is wrong (the MCP layer turns it into a tool error)."""


class JobNotFound(LookupError):
    """No such job (never existed, or deleted after the retention time)."""


class QueueFull(RuntimeError):
    """max_queued_jobs jobs are waiting already."""


class ReadFailed(RuntimeError):
    def __init__(self, kind: str, message: str) -> None:
        super().__init__(message)
        self.kind = kind


class FileRecord(BaseModel):
    name: str
    origin: Literal["upload", "output"]
    size: int
    sha256: str


class Job(BaseModel):
    """The record of a job (job.json)."""

    id: str
    request: str
    confirm: bool = False
    active_file: str = ""
    continue_from: str | None = None
    chain: list[str] = Field(default_factory=list, description="Earlier jobs, oldest first.")
    status: JobStatus = "queued"
    inputs: list[FileRecord] = Field(default_factory=list, description="Files of this job's uploads/.")
    outputs: list[FileRecord] = Field(default_factory=list)
    reply: str = ""
    failure: str | None = Field(default=None, description="agent_error, timed_out or interrupted.")
    message: str = ""
    verification: list[dict[str, Any]] = Field(default_factory=list)
    step_count: int = 0
    created_at: datetime
    started_at: datetime | None = None
    finished_at: datetime | None = None

    @property
    def finished(self) -> bool:
        return self.status in FINISHED

    @property
    def total_ms(self) -> float:
        if not self.started_at:
            return 0.0
        end = self.finished_at or datetime.now(UTC)
        return round((end - self.started_at).total_seconds() * 1000.0, 1)


class Progress(BaseModel):
    """What a running job is doing (from its stream-json events)."""

    step_count: int = 0
    current_activity: str = ""


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def mime_type(name: str) -> str:
    return mimetypes.guess_type(name)[0] or "application/octet-stream"


def _activity(step: dict[str, Any]) -> str:
    kinds = {"shell": "running", "read": "reading", "write": "writing", "edit": "editing",
             "skill": "skill", "search": "searching", "text": "", "tool": "tool"}
    prefix = kinds.get(step.get("k", ""), "")
    label = " ".join(str(step.get("note") or step.get("label") or "").split())
    return conv.clip(f"{prefix}: {label}" if prefix else label, 200)


Runner = Callable[[list[str], Path, Path, float], Awaitable[tuple[int | None, bool]]]


async def docker_runner(cmd: list[str], stdout: Path, stderr: Path, timeout: float) -> tuple[int | None, bool]:
    """Run the agent container; (exit code, timed out). The caller kills it on timeout."""
    with stdout.open("wb") as out, stderr.open("wb") as err:
        proc = await asyncio.create_subprocess_exec(*cmd, stdout=out, stderr=err)
        try:
            async with asyncio.timeout(timeout):
                return await proc.wait(), False
        except TimeoutError:
            return None, True
        except asyncio.CancelledError:
            proc.kill()
            await proc.wait()
            raise
        finally:
            if proc.returncode is None:
                proc.kill()
                await proc.wait()


class JobService:
    """Creates, runs, waits for and cancels jobs; keeps their records on disk."""

    def __init__(self, settings: Settings, *, runner: Runner | None = None, docker: Callable[..., Awaitable[Any]] | None = None) -> None:
        self.settings = settings
        self.root = settings.jobs_dir.resolve()
        self.jobs: dict[str, Job] = {}
        self._runner = runner or docker_runner
        self._docker = docker or (lambda *args, seconds=30: run_command([settings.docker_bin, *args], seconds))
        self._done: dict[str, asyncio.Event] = {}
        self._cancelled: set[str] = set()
        self._tasks: set[asyncio.Task[None]] = set()
        self._slots: asyncio.Semaphore | None = None
        self._live: dict[str, tuple[tuple[int, int], Progress]] = {}
        self._stopping = False
        self._draining = False
        self._janitor: asyncio.Task[None] | None = None

    # ---- lifecycle ---------------------------------------------------------
    async def start(self) -> None:
        """Load the records; jobs cut off by a restart become interrupted; old jobs go."""
        self.root.mkdir(parents=True, exist_ok=True)
        for record in sorted(self.root.glob("*/job.json")):
            try:
                job = Job.model_validate_json(record.read_text(encoding="utf-8"))
            except ValueError:
                log.warning("jobs.bad_record", path=str(record))
                continue
            if not job.finished:
                await self._docker("rm", "-f", self.container(job.id), seconds=30)
                self._finish(job, "interrupted", failure="interrupted",
                             message="The service was restarted while the job ran.")
            self.jobs[job.id] = job
        self.cleanup()
        self._janitor = asyncio.create_task(self._clean_every_hour())

    def drain(self) -> None:
        """The service is stopping: waiting calls return now with the state they have."""
        self._draining = True

    async def close(self) -> None:
        """Stop: running and queued jobs are stopped and recorded as interrupted."""
        self._stopping = True
        if self._janitor:
            self._janitor.cancel()
        for job in list(self.jobs.values()):
            if job.status == "running":
                await self._docker("kill", self.container(job.id), seconds=30)
        for task in list(self._tasks):
            task.cancel()
        await asyncio.gather(*self._tasks, return_exceptions=True)
        for job in self.jobs.values():
            if not job.finished:
                self._finish(job, "interrupted", failure="interrupted", message="The service was stopped while the job ran.")

    async def _clean_every_hour(self) -> None:
        while True:
            await asyncio.sleep(3600)
            self.cleanup()

    def cleanup(self) -> int:
        """Delete finished jobs older than job_retention_hours; how many went."""
        limit = datetime.now(UTC) - timedelta(hours=self.settings.job_retention_hours)
        gone = 0
        for job in list(self.jobs.values()):
            if job.finished and job.finished_at and job.finished_at < limit:
                shutil.rmtree(self.dir(job.id), ignore_errors=True)
                del self.jobs[job.id]
                gone += 1
        if gone:
            log.info("jobs.cleaned", deleted=gone)
        return gone

    # ---- paths -------------------------------------------------------------
    @staticmethod
    def container(job_id: str) -> str:
        return CONTAINER_PREFIX + job_id

    def dir(self, job_id: str) -> Path:
        return self.root / job_id

    def ws(self, job_id: str) -> Path:
        return self.dir(job_id) / "ws"

    def host_ws(self, job_id: str) -> Path:
        """The workspace as the docker daemon sees it (bind mounts)."""
        return self.settings.jobs_mount_dir / job_id / "ws"

    def save(self, job: Job) -> None:
        path = self.dir(job.id) / "job.json"
        tmp = path.with_suffix(".tmp")
        tmp.write_text(job.model_dump_json(indent=1), encoding="utf-8")
        tmp.replace(path)

    # ---- queries -----------------------------------------------------------
    def get(self, job_id: str) -> Job:
        job = self.jobs.get(str(job_id).strip())
        if job is None:
            raise JobNotFound(job_id)
        return job

    def counts(self) -> dict[str, int]:
        running = sum(1 for j in self.jobs.values() if j.status == "running")
        queued = sum(1 for j in self.jobs.values() if j.status == "queued")
        return {"running": running, "queued": queued}

    def queue_position(self, job: Job) -> int | None:
        if job.status != "queued":
            return None
        waiting = sorted((j for j in self.jobs.values() if j.status == "queued"), key=lambda j: j.created_at)
        return next(i for i, j in enumerate(waiting, start=1) if j.id == job.id)

    def progress(self, job: Job) -> Progress:
        if job.status != "running":
            return Progress(step_count=job.step_count)
        path = self.ws(job.id) / conv.EVENTS
        try:
            st = path.stat()
        except FileNotFoundError:
            return Progress()
        sig = (st.st_size, st.st_mtime_ns)
        cached = self._live.get(job.id)
        if not cached or cached[0] != sig:
            steps, _ = conv.summarize(conv.read_events(path))
            cached = self._live[job.id] = (sig, Progress(
                step_count=len(steps), current_activity=_activity(steps[-1]) if steps else ""))
        return cached[1]

    def output_path(self, job: Job, name: str) -> Path:
        if name not in {f.name for f in job.outputs}:
            raise JobNotFound(f"{job.id}/{name}")
        path = self.ws(job.id) / "outputs" / name
        if not path.is_file():
            raise JobNotFound(f"{job.id}/{name}")
        return path

    # ---- create ------------------------------------------------------------
    def chain_of(self, parent: Job) -> list[Job]:
        """The conversation that ends with ``parent``, oldest first."""
        chain = [self.get(job_id) for job_id in parent.chain]
        return [*chain, parent]

    def submit(
        self,
        request: str,
        files: list[tuple[str, bytes]] | None = None,
        *,
        continue_from: str = "",
        active_file: str = "",
        confirm: bool = False,
    ) -> Job:
        """Create a job and queue it; raises InvalidInput, JobNotFound or QueueFull."""
        request = request.strip()
        if not request:
            raise InvalidInput("request must not be empty")
        files = files or []
        s = self.settings
        if len(files) > s.max_input_files:
            raise InvalidInput(f"at most {s.max_input_files} files per job")
        total = 0
        for name, data in files:
            ext = Path(conv.clean_name(name)).suffix.lower()
            if ext not in conv.UPLOADABLE:
                raise InvalidInput(f"unsupported file type: {ext or '(no extension)'} ({name})")
            if not data:
                raise InvalidInput(f"empty file: {name}")
            if len(data) > s.max_input_file_bytes:
                raise InvalidInput(f"{name} is larger than {s.max_input_file_bytes} bytes")
            total += len(data)
        if total > s.max_input_total_bytes:
            raise InvalidInput(f"the files of a job may not exceed {s.max_input_total_bytes} bytes")
        parent = self.get(continue_from) if continue_from.strip() else None
        if parent is not None and not parent.finished:
            raise InvalidInput("continue_from: that job has not finished yet")
        if self.counts()["queued"] >= s.max_queued_jobs + max(0, s.max_running_jobs - self.counts()["running"]):
            raise QueueFull()

        job = Job(
            id=uuid.uuid4().hex,
            request=request,
            confirm=confirm,
            active_file=active_file.strip(),
            continue_from=parent.id if parent else None,
            chain=[*parent.chain, parent.id] if parent else [],
            created_at=datetime.now(UTC),
        )
        ws = self.ws(job.id)
        for sub in ("uploads", "scratch", "outputs"):
            (ws / sub).mkdir(parents=True, exist_ok=True)
        taken: set[str] = set()
        if parent is not None:  # the chain's files: every upload and deliverable so far
            for earlier, record in self._chain_files(parent):
                name = conv.unique_name(taken, record.name)
                taken.add(name)
                src = self.ws(earlier.id) / ("uploads" if record.origin == "upload" else "outputs") / record.name
                shutil.copyfile(src, ws / "uploads" / name)
        for name, data in files:
            name = conv.unique_name(taken, conv.clean_name(name))
            taken.add(name)
            (ws / "uploads" / name).write_bytes(data)
            job.inputs.append(FileRecord(name=name, origin="upload", size=len(data),
                                         sha256=hashlib.sha256(data).hexdigest()))
        if parent is not None and not job.active_file and parent.outputs:
            # "make the title red" without a file name means the file just delivered
            job.active_file = self._chain_names(parent)[-1]["name"]
        self.jobs[job.id] = job
        self._done[job.id] = asyncio.Event()
        self.save(job)
        log.info("jobs.submitted", job=job.id, files=len(files), continue_from=job.continue_from,
                 confirm=confirm, **request_field(s, request))
        task = asyncio.create_task(self._run(job))
        self._tasks.add(task)
        task.add_done_callback(self._tasks.discard)
        return job

    def _chain_files(self, parent: Job) -> list[tuple[Job, FileRecord]]:
        """Files of the chain in the order the conversation got them: each job's own
        inputs (job.inputs never lists the files it inherited), then its outputs."""
        return [(job, record) for job in self.chain_of(parent) for record in (*job.inputs, *job.outputs)]

    def _chain_names(self, parent: Job) -> list[dict[str, Any]]:
        """The conversation's files as conversation.build_prompt wants them (turn = 1-based)."""
        turn_of = {j.id: n for n, j in enumerate(self.chain_of(parent), start=1)}
        files, taken = [], set()
        for earlier, record in self._chain_files(parent):
            name = conv.unique_name(taken, record.name)
            taken.add(name)
            files.append({"name": name, "origin": record.origin, "turn": turn_of[earlier.id]})
        return files

    def prompt(self, job: Job) -> tuple[str, dict[int, list[str]], dict[int, str]]:
        """(prompt, work entries by turn, job id by turn) for the job's container."""
        turns: list[dict[str, Any]] = []
        files: list[dict[str, Any]] = []
        work: dict[int, list[str]] = {}
        ids: dict[int, str] = {}
        n = 1
        if job.continue_from:
            parent = self.get(job.continue_from)
            for n, earlier in enumerate(self.chain_of(parent), start=1):
                turns.append({"n": n, "text": earlier.request, "status": "done" if earlier.status == "done" else earlier.status,
                              "reply": earlier.reply, "outputs": [f.name for f in earlier.outputs],
                              "error": "" if earlier.status == "done" else (earlier.message or earlier.status)})
                ids[n] = earlier.id
                found = conv.entries(self.ws(earlier.id) / "scratch")
                if found:
                    work[n] = found
            files = self._chain_names(parent)
            n += 1
        taken = {f["name"] for f in files}
        for record in job.inputs:
            if record.name not in taken:
                files.append({"name": record.name, "origin": "upload", "turn": n})
        # Consent comes from `confirm` only: a "--onay" in the text may have been written by the
        # host's model (copied from a reply), not agreed to by the user.
        turn = {"n": n, "text": conv.no_flags(job.request), "status": "running", "confirm": job.confirm,
                "active": job.active_file if job.active_file in {f["name"] for f in files} else None}
        return conv.build_prompt({"turns": turns, "files": files}, turn, work), work, ids

    # ---- run ---------------------------------------------------------------
    def command(self, job: Job, prompt: str, work: dict[int, list[str]], ids: dict[int, str]) -> list[str]:
        s = self.settings
        host = self.host_ws(job.id)
        cmd = [s.docker_bin, "run", "--rm", "--name", self.container(job.id), "--label", f"{JOB_LABEL}={job.id}",
               "--network", s.docker_network, "--env-file", str(s.agent_env_file),
               "-e", f"GOTENBERG_URL={s.gotenberg_url}", "-e", "DOCGEN_OUTPUT_FORMAT=stream-json",
               "-e", f"DOCGEN_OUT_FILE=/workspace/{conv.EVENTS}",
               "-v", f"{host}:/workspace", "-v", f"{host / 'uploads'}:/workspace/uploads:ro"]
        for n in work:  # read-only: a follow-up must not change what an earlier job left behind
            cmd += ["-v", f"{self.host_ws(ids[n]) / 'scratch'}:{conv.PREVIOUS}/{n}:ro"]
        if s.job_memory_limit:
            cmd += ["--memory", s.job_memory_limit]
        if s.job_cpus:
            cmd += ["--cpus", s.job_cpus]
        return [*cmd, s.agent_image, prompt]

    async def _run(self, job: Job) -> None:
        try:
            if self._slots is None:
                self._slots = asyncio.Semaphore(self.settings.max_running_jobs)
            async with self._slots:
                if job.finished:  # cancelled while it was queued
                    return
                if job.id in self._cancelled or self._stopping:
                    self._finish(job, "cancelled" if job.id in self._cancelled else "interrupted",
                                 failure=None if job.id in self._cancelled else "interrupted")
                    return
                await self._execute(job)
        except asyncio.CancelledError:
            if not job.finished:
                self._finish(job, "interrupted", failure="interrupted", message="The service was stopped while the job ran.")
            raise
        except Exception as exc:  # never leave a job running
            log.exception("jobs.crashed", job=job.id)
            if not job.finished:
                self._finish(job, "failed", failure="agent_error", message=f"service error: {exc.__class__.__name__}")
        finally:
            self._live.pop(job.id, None)
            event = self._done.get(job.id)
            if event:
                event.set()

    async def _execute(self, job: Job) -> None:
        prompt, work, ids = self.prompt(job)
        jdir, ws = self.dir(job.id), self.ws(job.id)
        (jdir / "prompt.txt").write_text(prompt, encoding="utf-8")
        for path in [jdir, *jdir.rglob("*")]:  # the agent runs as uid 1000
            try:
                os.chmod(path, 0o777 if path.is_dir() else 0o666)
            except OSError:
                pass
        job.status, job.started_at = "running", datetime.now(UTC)
        self.save(job)
        log.info("jobs.started", job=job.id, chain=len(job.chain), previous=sorted(work))
        rc, timed_out = await self._runner(self.command(job, prompt, work, ids), jdir / "stdout.txt",
                                           jdir / "stderr.txt", float(self.settings.job_timeout_seconds))
        if timed_out or job.id in self._cancelled:
            await self._docker("kill", self.container(job.id), seconds=30)
        self._collect(job, rc, timed_out)

    def _collect(self, job: Job, rc: int | None, timed_out: bool) -> None:
        ws, jdir = self.ws(job.id), self.dir(job.id)
        stdout = (jdir / "stdout.txt").read_text(errors="replace") if (jdir / "stdout.txt").exists() else ""
        steps, result = conv.summarize(conv.read_events(ws / conv.EVENTS) or conv.events_from_stdout(stdout))
        (jdir / "steps.json").write_text(json.dumps(steps, ensure_ascii=False), encoding="utf-8")
        job.step_count = len(steps)
        job.reply = (result or {}).get("reply", "")
        outputs = ws / "outputs"
        job.outputs = [
            FileRecord(name=p.name, origin="output", size=p.stat().st_size, sha256=sha256(p))
            for p in sorted(outputs.iterdir()) if outputs.is_dir()
            if p.is_file() and p.suffix.lower() in conv.DELIVERABLE and not p.name.startswith(".")
        ] if outputs.is_dir() else []
        try:
            manifest = json.loads((outputs / MANIFEST).read_text(encoding="utf-8"))
            job.verification = [m for m in manifest if isinstance(m, dict)] if isinstance(manifest, list) else []
        except (OSError, ValueError):
            job.verification = []
        if job.id in self._cancelled:
            self._finish(job, "cancelled")
        elif self._stopping:  # killed by close(), not by the agent
            self._finish(job, "interrupted", failure="interrupted", message="The service was stopped while the job ran.")
        elif timed_out:
            self._finish(job, "timed_out", failure="timed_out",
                         message=f"The job did not finish within {self.settings.job_timeout_seconds} s.")
        elif rc != 0 or not result or result["is_error"]:
            stderr = [x for x in (jdir / "stderr.txt").read_text(errors="replace").splitlines()
                      if x.strip() and "running headless with --yolo" not in x] if (jdir / "stderr.txt").exists() else []
            why = (result or {}).get("error") or (stderr[-1] if stderr else f"agent exited with code {rc}")
            self._finish(job, "failed", failure="agent_error", message=conv.clip(why, 300))
        else:
            self._finish(job, "done")

    def _finish(self, job: Job, status: JobStatus, *, failure: str | None = None, message: str = "") -> None:
        job.status, job.failure, job.message = status, failure, message
        job.finished_at = datetime.now(UTC)
        self.save(job)
        log.info("jobs.finished", job=job.id, status=status, failure=failure, outputs=len(job.outputs),
                 steps=job.step_count, total_ms=job.total_ms)

    # ---- wait / cancel -----------------------------------------------------
    async def wait(self, job: Job, seconds: float,
                   on_progress: Callable[[Progress], Awaitable[None]] | None = None) -> Job:
        """Wait until the job finishes, at most ``seconds``; reports progress meanwhile."""
        event = self._done.get(job.id)
        deadline = time.monotonic() + max(0.0, seconds)
        while event is not None and not job.finished and not (self._stopping or self._draining):
            left = deadline - time.monotonic()
            if left <= 0:
                break
            try:
                async with asyncio.timeout(min(PROGRESS_SECONDS, left)):
                    await event.wait()
            except TimeoutError:
                pass
            if on_progress and not job.finished:
                await on_progress(self.progress(job))
        return job

    async def cancel(self, job: Job) -> Job:
        if job.finished:
            return job
        self._cancelled.add(job.id)
        if job.status == "queued":
            self._finish(job, "cancelled")
            event = self._done.get(job.id)
            if event:
                event.set()
            return job
        await self._docker("kill", self.container(job.id), seconds=30)
        log.info("jobs.cancel_requested", job=job.id)
        return job

    # ---- read --------------------------------------------------------------
    async def read(self, name: str, data: bytes | None = None, *, job: Job | None = None,
                   tables: bool = False, formatting: bool = False) -> tuple[str, bool]:
        """Text of a document (readback.py in a container without network); (text, truncated)."""
        clean = conv.clean_name(name)
        ext = Path(clean).suffix.lower()
        if ext not in READABLE:
            raise ReadFailed("unsupported_type", f"cannot read {ext or '(no extension)'} files")
        if formatting and ext not in FORMAT_READABLE:
            raise ReadFailed("unsupported_type", "formatting details exist for .docx, .pptx, .xlsx and .pdf only")
        tmp = self.root / "_reads" / uuid.uuid4().hex
        tmp.mkdir(parents=True)
        try:
            if data is not None:
                if len(data) > self.settings.max_input_file_bytes:
                    raise InvalidInput(f"{name} is larger than {self.settings.max_input_file_bytes} bytes")
                (tmp / clean).write_bytes(data)
            else:
                assert job is not None
                shutil.copyfile(self.output_path(job, name), tmp / clean)
            os.chmod(tmp, 0o755)
            os.chmod(tmp / clean, 0o644)
            host = self.settings.jobs_mount_dir / "_reads" / tmp.name
            args = ["run", "--rm", "--network", "none", "-v", f"{host}:/in:ro", "--entrypoint", "python",
                    self.settings.agent_image, "/opt/docgen/scripts/readback.py", f"/in/{clean}",
                    "--max-chars", str(self.settings.read_max_chars)]
            args += ["--tables"] if tables and ext == ".pdf" else []
            args += ["--format"] if formatting else []
            code, out = await self._docker(*args, seconds=self.settings.read_timeout_seconds)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        if code is None:
            raise ReadFailed("timeout", out)
        if code != 0:
            raise ReadFailed("read_failed", conv.clip(out.splitlines()[-1] if out else f"exit code {code}", 300))
        truncated = "\n... [truncated at " in out
        return out, truncated
