"""LibreOffice (headless) helpers: format conversion and formula recalculation.

LibreOffice runs inside the agent container with a private profile that
forces "always recalculate on load" for OOXML/ODF files.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from . import DocgenError

PROFILE = Path(tempfile.gettempdir()) / f"docgen-lo-profile-{os.getuid()}"
_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "docx": "MS Word 2007 XML",
           "pptx": "Impress MS PowerPoint 2007 XML", "pdf": None}


def soffice():
    for name in ("soffice", "libreoffice", "/usr/lib/libreoffice/program/soffice",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        path = shutil.which(name) or (name if os.path.exists(name) else None)
        if path:
            return path
    return None


def profile_url():
    user = PROFILE / "user"
    user.mkdir(parents=True, exist_ok=True)
    xcu = user / "registrymodifications.xcu"
    if not xcu.exists():
        xcu.write_text(_XCU)
    return PROFILE.as_uri()


def require():
    exe = soffice()
    if not exe:
        raise DocgenError("LibreOffice (soffice) is not installed in this environment")
    return exe


def convert(src, fmt, outdir=None, timeout=180):
    """Convert ``src`` to ``fmt`` (xlsx/docx/pptx/pdf) with LibreOffice;
    formulas are recalculated on load. Returns the output path."""
    exe = require()
    outdir = Path(outdir or tempfile.mkdtemp(prefix="docgen-lo-"))
    outdir.mkdir(parents=True, exist_ok=True)
    target = fmt if FILTERS.get(fmt) is None else f'{fmt}:{FILTERS[fmt]}'
    res = subprocess.run([exe, "--headless", "--norestore", "--nologo", f"-env:UserInstallation={profile_url()}",
                          "--convert-to", target, "--outdir", str(outdir), str(src)],
                         capture_output=True, text=True, timeout=timeout)
    out = outdir / (Path(src).stem + "." + fmt)
    if res.returncode != 0 or not out.exists():
        raise DocgenError(f"LibreOffice conversion failed: {res.stderr.strip() or res.stdout.strip()}")
    return out
