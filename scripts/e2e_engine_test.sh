#!/bin/bash
# DOCGEN end-to-end engine test (model bağımsız)
set -e
cd /workspace/scratch

echo "=== XLSX: create -> validate -> readback -> recalc ==="
python3 - <<'EOF'
from office_oxide import XlsxWriter
w = XlsxWriter()
s = w.add_sheet("Finans")
w.set_cell(s, 0, 0, "Ürün");   w.set_cell(s, 0, 1, "Miktar"); w.set_cell(s, 0, 2, "Fiyat")
w.set_cell_styled(s, 0, 0, "Ürün", bold=True, bg_color="#1F4E79")
w.set_cell(s, 1, 0, "Kalem");  w.set_cell(s, 1, 1, 3); w.set_cell(s, 1, 2, 10)
w.set_cell(s, 2, 0, "Defter"); w.set_cell(s, 2, 1, 7); w.set_cell(s, 2, 2, 25)
w.set_cell(s, 3, 0, "Toplam"); w.set_cell_styled(s, 3, 0, "Toplam", bold=True)
w.set_formula(s, 3, 1, "=SUM(B2:B3)")
w.set_column_width(s, 0, 18)
w.save("/workspace/outputs/test-finans.xlsx")
print("xlsx created")
EOF
python3 /opt/docgen/scripts/validate.py /workspace/outputs/test-finans.xlsx
python3 /opt/docgen/scripts/readback.py /workspace/outputs/test-finans.xlsx --tables | head -12
python3 /opt/docgen/scripts/recalc_check.py /workspace/outputs/test-finans.xlsx --expect "Miktar:10" --json
python3 /opt/docgen/scripts/render_check.py /workspace/outputs/test-finans.xlsx --expect-text "Toplam"

echo "=== RENDER CHECK: bozuk dosya FAIL vermeli ==="
# curl -s + ls bunu geçirirdi: Gotenberg'in hata metni "PDF" olarak kaydedilir
head -c 2000 /workspace/outputs/test-finans.xlsx > broken.xlsx
if python3 /opt/docgen/scripts/render_check.py broken.xlsx; then
  echo "HATA: render_check bozuk dosyayı geçirdi" >&2; exit 1
fi

echo "=== PDF: HTML -> chromium -> validate -> readback ==="
cat > page.html <<'EOF'
<!doctype html><html><head><meta charset="utf-8"><title>Test</title>
<style>
@page { size: A4; margin: 2cm; }
body { font-family: Calibri, sans-serif; color: #1A1A1A; }
h1 { color: #1F4E79; font-size: 24pt; }
h2 { color: #1F4E79; font-size: 16pt; margin-top: 18pt; }
table { border-collapse: collapse; width: 100%; }
th { background: #1F4E79; color: #fff; padding: 6px; }
td { border: 1px solid #BDD7EE; padding: 6px; }
</style></head><body>
<h1>DOCGEN PDF Test</h1>
<p>Temalı HTML-&gt;PDF akışı (Chromium). Tema: professional-blue.</p>
<h2>Özet Tablo</h2>
<table><tr><th>Kalem</th><th>Değer</th></tr><tr><td>Kalem</td><td>10</td></tr><tr><td>Defter</td><td>25</td></tr></table>
</body></html>
EOF
curl -s -f -F "files=@page.html;filename=index.html" "$GOTENBERG_URL/forms/chromium/convert/html" -o /workspace/outputs/test-rapor.pdf
python3 /opt/docgen/scripts/validate.py /workspace/outputs/test-rapor.pdf
python3 /opt/docgen/scripts/readback.py /workspace/outputs/test-rapor.pdf | head -8
echo "=== E2E OK ==="
