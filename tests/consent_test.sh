#!/bin/bash
# DOCGEN onay (--onay) kapısı testi — host'ta, docker ile; model gerekmez:
#   bash tests/consent_test.sh [imaj]
# Ajan sandbox'ında kütüphane force=True'yu yalnızca kullanıcının isteğinde
# --onay varken kabul eder. Onayı entrypoint, ajan başlamadan PID 1'in
# ortamına (DOCGEN_ONAY) sabitler. Test, ajanın yapabileceği denemeleri taklit
# eder: kendi ortamına DOCGEN_ONAY=1 yazmak, force=True, betiğe --onay vermek.
set -euo pipefail
IMAGE="${1:-docgen-agent:latest}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
FAIL=0

cat > "$TMP/probe.py" <<'EOF'
import os, subprocess, sys
os.environ["DOCGEN_ONAY"] = "1"  # the agent's own environment must not count
from docgen import NeedsConfirmation, xlsx as X
from docgen.package import Package
try:
    X.formula_to_value(Package("/fx/pro_satis_paneli.xlsx"), "Özet!B4", force=True)
    lib = "allowed"
except NeedsConfirmation:
    lib = "refused"
r = subprocess.run([sys.executable, "/opt/docgen/scripts/xlsx_struct.py", "/fx/pro_satis_paneli.xlsx", "/tmp/o.xlsx",
                    "delete-rows", "--sheet", "Parametreler", "--at", "2", "--onay", "--no-recalc"],
                   capture_output=True, text=True)
print(f"PROBE lib={lib} script={'allowed' if r.returncode == 0 else 'refused'}")
EOF
# entrypoint'in çağırdığı "qwen" yerine probe; Gotenberg yerine bir dosya
mkdir -p "$TMP/bin"
printf '#!/bin/sh\nexec python /t/probe.py\n' > "$TMP/bin/qwen"
touch "$TMP/health"
chmod -R a+rX "$TMP"
chmod a+x "$TMP/bin/qwen"

check() {  # name, expected, got
  if [ "$3" = "$2" ]; then echo "ok    $1"; else echo "FAIL  $1: expected '$2', got '$3'"; FAIL=1; fi
}
mounts=(-v "$ROOT/evals/fixtures:/fx:ro" -v "$TMP:/t:ro")
as_pid1() {  # probe as the container's main process with the given environment
  docker run --rm --entrypoint env "${mounts[@]}" "$IMAGE" "$@" python /t/probe.py | grep '^PROBE' || true
}
via_entrypoint() {  # request, extra docker args...
  local req="$1"; shift
  docker run --rm "$@" "${mounts[@]}" -e MODEL_NAME=x -e MODEL_ENDPOINT=http://x -e GOTENBERG_URL=file:///t \
    -e PATH=/t/bin:/opt/docgen/node_modules/.bin:/usr/local/bin:/usr/bin:/bin "$IMAGE" "$req" | grep '^PROBE' || true
}

check "outside the sandbox force=True stands" "PROBE lib=allowed script=allowed" "$(as_pid1 X=1)"
check "PID 1 DOCGEN_ONAY=0 refuses" "PROBE lib=refused script=refused" "$(as_pid1 DOCGEN_ONAY=0)"
check "PID 1 DOCGEN_ONAY=1 allows" "PROBE lib=allowed script=allowed" "$(as_pid1 DOCGEN_ONAY=1)"
check "request without --onay refuses" "PROBE lib=refused script=refused" \
  "$(via_entrypoint "Özet sayfasındaki formülleri değere çevir.")"
check "host -e DOCGEN_ONAY=1 does not count" "PROBE lib=refused script=refused" \
  "$(via_entrypoint "Özet sayfasındaki formülleri değere çevir." -e DOCGEN_ONAY=1)"
check "'--onay' inside a word does not count" "PROBE lib=refused script=refused" \
  "$(via_entrypoint "x--onay ve --onaylı yaz")"
check "request with --onay allows" "PROBE lib=allowed script=allowed" \
  "$(via_entrypoint "Özet sayfasındaki formülleri değere çevir. --onay")"
rc=0
docker run --rm --init "${mounts[@]}" -e MODEL_NAME=x -e MODEL_ENDPOINT=http://x "$IMAGE" "x --onay" >/dev/null 2>&1 || rc=$?
check "entrypoint refuses to run when it is not PID 1" "2" "$rc"

[ "$FAIL" -eq 0 ] && echo "consent: all checks passed" || echo "consent: FAILED"
exit "$FAIL"
