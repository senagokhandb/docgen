#!/bin/sh
# Writes deploy/stack.env for docker compose (--env-file): the bearer token of the MCP
# server, the job directory as the docker daemon sees it and the group of the docker
# socket. Existing values are kept. The file must not be committed or copied into an image.
set -eu
cd "$(dirname "$0")"
FILE=stack.env
touch "$FILE"
chmod 600 "$FILE"
add() {
  if ! grep -q "^$1=" "$FILE"; then
    printf '%s=%s\n' "$1" "$2" >> "$FILE"
    echo "set $1"
  else
    echo "kept $1"
  fi
}
add DOCGEN_MCP_HTTP_TOKEN "$(openssl rand -hex 32)"
JOBS="$(cd .. && pwd)/data/mcp-jobs"
mkdir -p "$JOBS"
add DOCGEN_JOBS_HOST_DIR "$JOBS"
# The socket's group inside a container: the host's on Linux, root (0) on Docker Desktop.
if [ "$(uname -s)" = "Linux" ]; then GID="$(stat -c %g /var/run/docker.sock)"; else GID=0; fi
add DOCKER_SOCKET_GID "$GID"
echo "secrets are in deploy/$FILE (not shown)"
