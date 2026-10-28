#!/bin/bash
# ============================================================
# DOCGEN entrypoint
# 0) istekteki --onay'ı DOCGEN_ONAY olarak bu sürecin (PID 1) ortamına sabitler
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

# --- 0) kullanıcı onayı (--onay) ---
# Yıkıcı değişiklikler (formül -> değer, başka hücrelerin okuduğu satırı silme,
# bağlantı verilen slaytı silme ...) yalnızca istekte "--onay" varsa yapılır.
# Karar burada, ajan başlamadan verilir ve DOCGEN_ONAY=1|0 olarak bu sürecin
# (PID 1) ortamına sabitlenir; docgen kütüphanesi onu /proc/1/environ'dan okur
# (docgen.require_consent). Ajan kendi süreçlerinde değişken atayabilir ama
# PID 1'in ortamını değiştiremez: force=True ya da DOCGEN_ONAY=1 yazması onay
# sayılmaz. .env / docker -e ile gelen DOCGEN_ONAY yok sayılır; tek kaynak istek.
if [ "$$" -ne 1 ]; then
  echo "[docgen] entrypoint konteynerin ana süreci (PID 1) olmalı; --init kullanmayın." >&2
  exit 2
fi
ONAY=0
if [[ " $REQUEST " =~ [[:space:]]--onay[[:space:]] ]]; then ONAY=1; fi
if [ "${DOCGEN_ONAY:-}" != "$ONAY" ]; then
  exec env DOCGEN_ONAY="$ONAY" "$0" "$@"
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
# DOCGEN_QWEN_ARGS: ek qwen bayrakları (varsayılan boş), ör. "--max-wall-time 15m".
# --auth-type openai zorunlu: Qwen Code 0.26.1 headless modda kimlik doğrulama
# türü seçilmeden çalışmaz ("No auth type is selected").
# JSON çıktısı önce dosyaya yazılır: Node, pipe'a yazarken süreç kapanırsa
# 64 KB'ı aşan çıktıyı keser (uzun işlerde yarım JSON + kayıp son yanıt).
# İstekte "@" varsa (e-posta adresi vb.) Qwen Code onu "@dosya" referansı
# sanıp önüne boşluk ekler ("ali@firma.com" -> "ali @firma.com"). Bu durumda
# istek bir dosyaya yazılır ve tek bir @referansla verilir: dosya içeriği
# modele olduğu gibi gider, içindeki "@" ayrıştırılmaz.
# DOCGEN_OUTPUT_FORMAT: json (varsayılan, iş bitince tek JSON dizisi) |
# stream-json (satır başına bir olay, iş sürerken yazılır).
# DOCGEN_OUT_FILE: çıktının yazılacağı dosya (varsayılan: geçici dosya). Verilirse
# silinmez — sohbet arayüzü (ui/) /workspace altındaki bu dosyayı konteyner
# çalışırken okuyup ajanın adımlarını canlı gösterir.
PROMPT="$REQUEST"
REQ_FILE=""
if [[ "$REQUEST" == *@* ]]; then
  mkdir -p /workspace/scratch
  REQ_FILE="/workspace/scratch/docgen-request.txt"
  printf '%s\n' "$REQUEST" > "$REQ_FILE"
  PROMPT="The user's request is the exact content of @$REQ_FILE — carry out that request."
fi
OUT_FILE="${DOCGEN_OUT_FILE:-$(mktemp)}"
set +e
# shellcheck disable=SC2086
qwen "$PROMPT" \
  --auth-type openai \
  --model "${MODEL_NAME:?MODEL_NAME .env dosyasinda tanımlı olmali}" \
  --openai-base-url "${MODEL_ENDPOINT:?MODEL_ENDPOINT .env dosyasinda tanımlı olmali}" \
  --openai-api-key "${MODEL_API_KEY:-no-key}" \
  --yolo \
  -o "${DOCGEN_OUTPUT_FORMAT:-json}" ${DOCGEN_QWEN_ARGS:-} > "$OUT_FILE"
RC=$?
set -e
cat "$OUT_FILE"
if [ -z "${DOCGEN_OUT_FILE:-}" ]; then rm -f "$OUT_FILE"; fi
rm -f ${REQ_FILE:+"$REQ_FILE"}
exit "$RC"
