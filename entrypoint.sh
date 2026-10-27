#!/bin/bash
# ============================================================
# DOCGEN entrypoint
# 1) skills + sistem kurallarını agent evine kopyalar
# 2) Gotenberg'in hazır olmasını bekler
# 3) Qwen Code'u tek seferlik (one-shot) çalıştırır:
#    - model:  .env'den MODEL_NAME
#    - endpoint: .env'den MODEL_ENDPOINT (OpenAI-uyumlu)
#    - --yolo: denetimsiz mod (sandbox içinde zaten non-root)
#    - -o json: çıktı JSON
# ============================================================
set -euo pipefail

REQUEST="${1:-}"
if [ -z "$REQUEST" ]; then
  echo "Kullanım: docker compose run --rm agent \"<belge istegi>\"" >&2
  echo "Örnek:   docker compose run --rm agent \"Q3 finans raporu.docx oluştur\"" >&2
  exit 1
fi

# --- 1) agent evi: skills ve sistem kuralları ---
mkdir -p /home/agent/.qwen/skills
cp -rf /opt/docgen/skills/. /home/agent/.qwen/skills/
cp -f  /opt/docgen/QWEN.md   /home/agent/.qwen/QWEN.md

# --- 2) Gotenberg bekle (iç ağ) ---
GOTENBERG_URL="${GOTENBERG_URL:-http://gotenberg:3000}"
for i in $(seq 1 30); do
  if curl -sf "$GOTENBERG_URL/health" >/dev/null 2>&1; then
    echo "[docgen] Gotenberg hazır: $GOTENBERG_URL"
    break
  fi
  if [ "$i" -eq 30 ]; then
    echo "[docgen] UYARI: Gotenberg 60 sn içinde yanıt vermedi, devam ediliyor." >&2
  fi
  sleep 2
done

# --- 3) agent: Qwen Code one-shot ---
exec qwen "$REQUEST" \
  --model "${MODEL_NAME:?MODEL_NAME .env dosyasinda tanımlı olmali}" \
  --openai-base-url "${MODEL_ENDPOINT:?MODEL_ENDPOINT .env dosyasinda tanımlı olmali}" \
  --openai-api-key "${MODEL_API_KEY:-no-key}" \
  --yolo \
  -o json
