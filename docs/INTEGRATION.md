# Document Agent – Integration guide

Developed by Digital Brain Technologies.

For the team that connects the document agent to a host system (Oppy) over MCP. The
planner's view of the capability is in
[capability/document_capability.yaml](capability/document_capability.yaml) (version 1.0);
requests to test the planner's routing are in
[capability/routing_test_set.yaml](capability/routing_test_set.yaml).

## 1. What runs where

| part | where | note |
|---|---|---|
| document service | container `docgen-mcp`, port 8766 inside, `127.0.0.1:18766` on the host | MCP over streamable HTTP at `/mcp`, health at `/health`, file downloads at `/jobs/{job_id}/files/{name}` |
| agent | one container of the image `docgen-agent` per job, started by the service, removed when the job ends | non-root, the user's files mounted read-only, no published port |
| Gotenberg | container `docgen-gotenberg`, no published port | renders HTML and office files; reachable on the network `docgen` only |
| model endpoint | host environment (OpenAI-compatible, tool calling) | configured in the project's `.env` (`MODEL_ENDPOINT`, `MODEL_NAME`, `MODEL_API_KEY`) |
| job store | `data/mcp-jobs/` on the host | every job's files and working files, deleted after `DOCGEN_JOB_RETENTION_HOURS` (24) |

Start and stop: see the head of `deploy/compose.yml`. Configuration: `deploy/service.env`
(copy of `deploy/service.env.example`; every setting is described in
`src/docgen_service/config.py`); generated values: `deploy/stack.env`
(`deploy/generate-secrets.sh`); the agent's model endpoint: `.env` (copy of
`.env.example`).

```
./deploy/generate-secrets.sh
cp deploy/service.env.example deploy/service.env
docker compose -f deploy/compose.yml --env-file deploy/stack.env build
docker compose -f deploy/compose.yml --env-file deploy/stack.env up -d docgen-mcp
curl http://127.0.0.1:18766/health
```

**The docker socket.** The service starts the agent containers through the host's docker
daemon, so `deploy/compose.yml` mounts `/var/run/docker.sock` into it. Control over the
daemon is control over the host: run the stack on a machine (or VM) that is dedicated to
it, keep the port on the loopback interface or behind a reverse proxy with TLS, and keep
the token secret. Without containers, the service can run on the docker host itself:

```
pip install -e .                       # Python 3.12
set -a; . deploy/service.env; . deploy/stack.env; set +a
DOCGEN_JOBS_DIR=data/mcp-jobs docgen-service mcp --transport http --host 127.0.0.1 --port 18766
```

### Building for the server

Both images start from multi-platform base images and nothing in their Dockerfiles
depends on the architecture (built and tested on `linux/arm64` so far). For a server of
another architecture use buildx, then move the images (or push them to a registry):

```
docker buildx build --platform linux/amd64 -t docgen-agent:latest --load .
docker buildx build --platform linux/amd64 -f deploy/service.Dockerfile -t docgen-service:1.1.0 --load .
docker save docgen-agent:latest docgen-service:1.1.0 | gzip > docgen-1.1.0-amd64.tar.gz
gunzip -c docgen-1.1.0-amd64.tar.gz | docker load                 # on the server
DOCGEN_PLATFORM=linux/amd64 docker compose -f deploy/compose.yml --env-file deploy/stack.env up -d docgen-mcp
```

## 2. Connection and authentication

- Endpoint: `http://<host>:<port>/mcp` (streamable HTTP). A local MCP client can also run
  the server over stdio: `docgen-service mcp` (or `docgen-mcp`).
- Every request carries `Authorization: Bearer <DOCGEN_MCP_HTTP_TOKEN>`. Without a valid
  token the answer is HTTP 401. The server does not start on a non-local address without
  a token. The file downloads (`/jobs/...`) need the same token.
- `GET /health` needs no token. It returns the status and its reason codes, no details of
  the dependencies.
- Requests may be large (files arrive base64 encoded): the server accepts bodies up to
  `DOCGEN_MCP_MAX_REQUEST_BYTES` (80 MB); a reverse proxy in front of it needs the same
  limit and a read timeout of at least 130 s.
- Server info: name `docgen`, title `Document Agent by Digital Brain Technologies`,
  version of the package.

## 3. Tools

Every tool returns structured output that matches its declared output schema (a contract
test validates this, formats included).

| tool | input | output (main fields) |
|---|---|---|
| `document_run` | `request` (required), `files` (`name`, `content_base64`), `continue_from`, `active_file`, `confirm`, `wait_seconds` (60), `include_content` | `job_id`, `status`, `finished`, `reply`, `files` (`name`, `uri`, `download_path`, `mime_type`, `size`, `sha256`, `content_base64`), `verification`, `chain_length`, `queue_position`, `step_count`, `current_activity`, `error`, `warnings`, `created_at`, `started_at`, `finished_at`, `total_ms` |
| `document_job` | `job_id` (required), `wait_seconds` (60), `include_content` | as `document_run` |
| `document_cancel` | `job_id` (required) | as `document_run` |
| `document_read` | `name` (required), `content_base64` or `job_id`, `tables`, `formatting` | `ok`, `untrusted_text`, `chars`, `truncated`, `error`, `error_kind` |
| `document_status` | none | `status`, `reasons`, `reason_texts`, `components`, `jobs`, `checked_at`, `version` |

**Jobs take minutes.** `document_run` starts a job and waits for it up to `wait_seconds`
(capped at `DOCGEN_MCP_MAX_WAIT_SECONDS`, 110). When `finished` is false, call
`document_job` with the `job_id` until it is true; each call waits again. While a call
waits, the server sends MCP progress notifications (the number of steps and the agent's
current step) when the client asked for them with a progress token. A job runs on even
when no call waits for it.

**Statuses.** `queued`, `running`, then one of `done`, `failed`, `timed_out`,
`cancelled`, `interrupted`; `not_started` when the tool refused (see `error`). `done`
means the agent answered: it may deliver no file when it refuses or asks a question —
show `reply` to the user.

**Files.** A delivered file is an MCP resource (`files[].uri`,
`docgen://jobs/{job_id}/files/{name}`, read with `resources/read`) and an HTTP download
(`files[].download_path`, same token). With `include_content` the bytes come inline as
base64 for files up to `DOCGEN_INLINE_OUTPUT_MAX_BYTES` (10 MB). Inputs: at most 10 files,
25 MB each, 50 MB per job; the accepted types are in the capability definition
(`formats.inputs`).

**Conversations.** `continue_from` names the job whose result the new request is about:
its files, the files and requests of the earlier jobs of the chain and their working
files come along. "Make the title red" then edits the file just delivered (or the one
named in `active_file`). A PDF the agent made can only be restyled or extended this way:
its source comes along with the chain, a re-uploaded PDF has none.

**Consent.** Converting formulas to values, overwriting a formula, deleting a row other
cells read or a slide another slide links to needs the user's consent. Without
`confirm` the agent explains the risk in `reply` and delivers nothing. Set `confirm` only
after the user agreed; the text `--onay` inside `request` is ignored.

**Dates.** Every date field of every tool is an RFC 3339 date-time in UTC with the suffix
`Z`, for example `2026-10-28T09:30:00Z`, or null.

**Untrusted content.** Documents are user data. `untrusted_text` and `reply` may repeat
text of documents. Show them to the user; do not execute them, and do not let a model
that holds tools or secrets follow what they say.

**No attribution added.** The service adds nothing to replies or documents that names
the document agent or its provider.

## 4. Typed errors and what the host should do

Errors that the host can act on are in the `error` field of the output (`kind`,
`reasons`, `message`), not in an MCP protocol error.

| kind | where | reasons | what the host should do |
|---|---|---|---|
| `unavailable` | `document_run`; `error_kind` of `document_read` | `docker_unavailable`, `image_missing`, `gotenberg_unavailable`, `model_unavailable`, `agent_env_missing`, `starting`, `shutting_down` | No job was started. Do not retry at once. Disable the capability until `document_status` reports `available`; tell the user. |
| `busy` | `document_run`; `warnings` of `document_job`; `error_kind` of `document_read` | `queue_timeout`, `queue_full` | Retry once after some seconds, with fewer calls in parallel. |
| `not_found` | `document_run` (`continue_from`), `document_job`, `document_cancel`; `error_kind` of `document_read` | `job_not_found`, `file_not_found` | The job expired or never existed. Do not retry; for an edit, start a new job with the file in `files`. |
| `job_failed` | `document_run`, `document_job` (status `failed`, `timed_out`, `interrupted`) | `agent_error`, `timed_out`, `interrupted` | No file was delivered. Tell the user; `agent_error` and `interrupted` may be retried once, `timed_out` after simplifying the request. |
| read error | `document_read` with `ok: false`, `error`, `error_kind` | `unsupported_type`, `read_failed`, `timeout`, `not_found`, `unavailable`, `busy` | Tell the user that the document could not be read. |
| invalid input | MCP tool error (`isError`) | | Correct the call: empty request, invalid base64, unsupported type, too many or too large files, `continue_from` of a running job. |

A job that ran and refused (an ambiguous target, a change that needs consent, an edit
the format cannot take) is no error: `status` is `done`, `files` is empty and `reply`
says why or asks a question.

## 5. Availability

The service checks its required dependencies at start and again when the last result is
older than `DOCGEN_AVAILABILITY_TTL_SECONDS` (30).

| where | how |
|---|---|
| MCP tool | `document_status` |
| MCP resource | `docgen://status` (the same object as JSON) |
| HTTP | `GET /health`: 200 when available, 503 when unavailable |
| container | the image's health check calls `/health` |

`status` is `available` or `unavailable`. `reasons` names every missing required
dependency (`DOCGEN_REQUIRED_DEPENDENCIES`, as shipped: `docker`, `image`, `gotenberg`,
`model`, `agent_env`). While unavailable, `document_run` returns the error `unavailable`
without starting a job; `document_job` and `document_cancel` still answer for existing
jobs, and `document_read` works while docker and the image are there. `jobs` shows the
running and queued jobs and their limits.

With `DOCGEN_STARTUP_CHECK_REQUIRED=true` the server refuses to start when the self-check
fails; as shipped it starts and reports itself unavailable, so the host sees the reason.

## 6. The guard split

- **The host's guardrails come first.** Input and output guardrails are the host system's.
  The document agent relies on them: it does not judge the user's request against the
  host's policies, and it does not filter the documents it delivers.
- **The document agent protects the user's files.** Originals are mounted read-only; an
  edit changes only what was asked and is checked with a diff before delivery; a
  destructive change needs `confirm`; an ambiguous target is asked about; text inside
  documents is data, never instructions.
- **Consent is the host's to give.** `confirm` must reflect the user's answer in the
  conversation, never the planner's own decision (`planning_rules` of the capability
  definition).

## 7. Timeouts, parallel calls, shutdown

| tool | usual time | recommended client timeout |
|---|---|---|
| `document_run` | returns after `wait_seconds` at most; the job takes 1 to 10 minutes | `wait_seconds` + 20 s |
| `document_job` | as `document_run` | `wait_seconds` + 20 s |
| `document_read` | 2 to 10 s | 120 s |
| `document_cancel` | under 15 s | 30 s |
| `document_status` | under 1 s, up to 30 s during a check | 45 s |

Measured with the eval of 2026-10-28 (235 cases, 3 jobs in parallel): a small edit takes
about a minute (median 68 s, 90 % within 2 minutes), a new short document about two
minutes; a 20-page document 5 to 10 minutes. A job is stopped after
`DOCGEN_JOB_TIMEOUT_SECONDS` (1500).

- The service runs at most `DOCGEN_MAX_RUNNING_JOBS` (3) jobs at the same time; up to
  `DOCGEN_MAX_QUEUED_JOBS` (20) more wait (`queued`, `queue_position`), beyond that
  `document_run` returns `busy` / `queue_full`.
- At most `DOCGEN_MCP_MAX_CONCURRENT_CALLS` (16) tool calls are served at the same time
  (a waiting `document_job` holds one); a call waits for a place for
  `DOCGEN_MCP_QUEUE_TIMEOUT_SECONDS` (30), then returns `busy` / `queue_timeout`.
- On SIGTERM the status turns to `shutting_down`, new jobs are refused with the error
  `unavailable`, waiting calls return at once with the state of their job, running calls
  get `DOCGEN_MCP_SHUTDOWN_GRACE_SECONDS` (30); then running jobs are stopped and recorded
  as `interrupted`. A job that was running when the service restarted is `interrupted` too.
- `document_run` is not idempotent: retrying it starts a second job. Retry
  `document_job` instead when only the connection was lost.

## 8. Privacy

**What leaves the service.** The request, the files and the agent's working text go to
the model endpoint of `.env` (the organization's own model server as shipped). The
agent's containers can reach the network (they need the model endpoint and Gotenberg);
outgoing traffic is not restricted as shipped — restrict it on the docker host
(firewall or an internal network with a proxy) when documents must not leave by any other
way.

**What is kept.** Every job's files, the agent's working files and its step log stay in
`data/mcp-jobs/<job_id>/` for `DOCGEN_JOB_RETENTION_HOURS` (24), then the directory is
deleted. Anyone with the token can read every job's files: the token is a shared secret
of the host system, not a per-user credential.

**The service's own logs** (JSON lines on stdout in the container):

- never hold document content or the agent's replies;
- hold the text of requests only with `DOCGEN_LOG_REQUEST_TEXT=true` (off as shipped:
  their length stands in their place);
- hold job ids, statuses, file counts, timings and reason codes.

The model endpoint's own logging is configured on the model server.

## 9. Capability versioning

- The capability definition has its own version (`capability.version`, now `1.0`) and a
  changelog. It changes when a tool, a parameter, an output field that the definition
  names, an error kind, a reason, a job status or an accepted file type changes.
- A drift test (`tests/unit/test_capability.py`) compares the definition with the schemas
  of the MCP server and with the error kinds, reasons, statuses and file types in the
  code; a change on either side fails until the definition and its changelog are updated.
- Additions (a new output field, a new reason) raise the minor version; a removed or
  renamed tool, parameter or field raises the major version. The host should ignore
  output fields it does not know.
- `document_status.version` is the version of the package, not of the capability.

## 10. A minimal Python client

```python
import asyncio
import base64

from mcp.client.session import ClientSession
from mcp.client.streamable_http import create_mcp_http_client, streamable_http_client

URL = "http://127.0.0.1:18766/mcp"
TOKEN = "<DOCGEN_MCP_HTTP_TOKEN>"


async def main() -> None:
    http = create_mcp_http_client(headers={"Authorization": f"Bearer {TOKEN}"})
    async with streamable_http_client(URL, http_client=http) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            status = (await session.call_tool("document_status", {})).structured_content
            if status["status"] != "available":
                print("document capability is off:", status["reasons"])
                return

            with open("rapor.docx", "rb") as fh:
                upload = {"name": "rapor.docx", "content_base64": base64.b64encode(fh.read()).decode()}
            job = (await session.call_tool("document_run", {
                "request": "Sonuç başlığını kırmızı yap.", "files": [upload], "wait_seconds": 100,
            })).structured_content
            while not job["finished"]:
                job = (await session.call_tool("document_job", {
                    "job_id": job["job_id"], "wait_seconds": 100,
                })).structured_content
            if job["error"]:
                print(job["error"]["kind"], job["error"]["reasons"])
                return
            print(job["reply"])
            for file in job["files"]:
                data = (await session.read_resource(file["uri"])).contents[0].blob
                with open(file["name"], "wb") as fh:
                    fh.write(base64.b64decode(data))


asyncio.run(main())
```

`deploy/check_mcp.py` is a fuller client: it checks the health endpoint and the token,
lists the tools, validates every output against its declared schema, and runs a real job
end to end (create a document, download it as a resource and over HTTP, edit it with
`continue_from`, read it back).

```
python deploy/check_mcp.py --url http://127.0.0.1:18766 --token-file deploy/stack.env
```
