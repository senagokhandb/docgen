# DOCGEN service image: the MCP server of the document agent. Developed by Digital Brain Technologies.
#
# Build (from the project root):   docker build -f deploy/service.Dockerfile -t docgen-service:<version> .
# For a server of another arch:    docker buildx build --platform linux/amd64 -f deploy/service.Dockerfile \
#                                      -t docgen-service:<version> --load .
# Holds the service package, its pinned dependencies and the docker client. The agent
# itself is the docgen-agent image (Dockerfile): the service starts one container of it
# per job through the host's docker daemon. No secret and no configuration of a
# deployment is in the image: both come from the environment at start
# (deploy/service.env, deploy/stack.env).

# Pinned by digest (python 3.12, Debian 13). Bump deliberately and rebuild.
FROM python:3.12-slim@sha256:78387bc3881b8273120a12ebe6c1ab22b018ccc2c9adf565ae1ac9b536e184ea

LABEL org.opencontainers.image.title="DOCGEN Service" \
      org.opencontainers.image.vendor="Digital Brain Technologies" \
      org.opencontainers.image.description="Document agent (DOCX, PPTX, XLSX, PDF) with an MCP server. Developed by Digital Brain Technologies."

# The docker client only (the daemon is the host's, through its socket).
COPY --from=docker:29.2.1-cli /usr/local/bin/docker /usr/local/bin/docker

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DOCGEN_LOG_FORMAT=json \
    DOCGEN_LOG_STREAM=stdout \
    DOCGEN_MCP_HTTP_HOST=0.0.0.0 \
    DOCGEN_MCP_HTTP_PORT=8766

WORKDIR /app

# Dependencies first (their layer changes less often than the code).
COPY deploy/requirements.lock /app/requirements.lock
RUN pip install -r /app/requirements.lock

COPY pyproject.toml README.md /app/
COPY src /app/src
RUN pip install --no-deps . \
    && useradd --uid 10001 --user-group --create-home --shell /usr/sbin/nologin docgen \
    && chmod -R a+rX /app

USER 10001:10001
ENV HOME=/home/docgen

EXPOSE 8766

# Healthy means available: the self-check found every required dependency (HTTP 200).
HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
    CMD ["python", "-c", "import sys, urllib.request; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8766/health', timeout=8).status == 200 else 1)"]

STOPSIGNAL SIGTERM
ENTRYPOINT ["docgen-service"]
CMD ["mcp", "--transport", "http"]
