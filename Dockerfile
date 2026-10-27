# ============================================================
# DOCGEN agent imajı
# İçerik:
#   - Python 3.12 + office_oxide / pdfplumber / lxml / openpyxl /
#     python-docx / python-pptx / pypdfium2
#   - LibreOffice (headless, nogui): formül yeniden hesabı (yalnızca hesap
#     motoru; dosyayı yeniden yazmaz) ve eski format (.doc/.xls/.ppt) dönüşümü
#   - Gotenberg ile aynı fontlar (Caladea, Carlito, Liberation, DejaVu):
#     PDF metin düzenlemede aynı aileyle yeniden çizim için
#   - Node 22 + Qwen Code CLI + docx-js + pptxgenjs
#   - docgen skills, sistem kuralları (QWEN.md), docgen kütüphanesi (lib/),
#     QA scriptleri, tema tokenleri, şablonlar, ECMA-376 XSD şemaları
#   - non-root "agent" kullanıcısı (uid 1000) — sandbox
# ============================================================
FROM python:3.12-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    NODE_PATH=/opt/docgen/node_modules \
    PYTHONPATH=/opt/docgen/lib \
    PATH=/opt/docgen/node_modules/.bin:$PATH

# --- sistem bağımlılıkları + LibreOffice (headless) + fontlar + Node 22 ---
RUN apt-get update && apt-get install -y --no-install-recommends \
        curl ca-certificates gnupg unzip zip \
        libreoffice-calc-nogui libreoffice-writer-nogui libreoffice-impress-nogui \
        fontconfig fonts-crosextra-caladea fonts-crosextra-carlito fonts-liberation2 fonts-dejavu-core \
    && curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

# --- Python motoru ---
COPY requirements.txt /opt/docgen/requirements.txt
RUN pip install --no-cache-dir -r /opt/docgen/requirements.txt

# --- Node: Qwen Code CLI + ofis üretimi kütüphaneleri ---
COPY package.json /opt/docgen/package.json
RUN npm install -g @qwen-code/qwen-code@0.26.1 \
    && cd /opt/docgen && npm install --omit=dev --no-audit --no-fund \
    && npm cache clean --force

# --- uygulama kodu ---
COPY entrypoint.sh /opt/docgen/entrypoint.sh
COPY QWEN.md       /opt/docgen/QWEN.md
COPY skills/       /opt/docgen/skills/
COPY lib/          /opt/docgen/lib/
COPY scripts/      /opt/docgen/scripts/
COPY themes/       /opt/docgen/themes/
COPY schemas/      /opt/docgen/schemas/
# host dosya izinleri (600) COPY ile korunur; imajda herkese okunur yap
RUN chmod -R a+rX /opt/docgen && chmod 755 /opt/docgen/entrypoint.sh

# --- tema şablonlarını imaj derleme sırasında üret (tekrarlanabilir) ---
RUN python /opt/docgen/scripts/build_templates.py

# --- non-root sandbox kullanıcısı ---
RUN useradd -m -u 1000 -s /bin/bash agent \
    && mkdir -p /workspace /home/agent/.qwen \
    && chown -R agent:agent /workspace /home/agent

USER agent
WORKDIR /workspace
ENTRYPOINT ["/opt/docgen/entrypoint.sh"]
