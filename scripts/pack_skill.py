#!/usr/bin/env python3
"""Pack a skill directory into a .skill archive (zip).

Usage:
    pack_skill.py <skill_dir> [-o out.skill]

The archive contains the skill folder at the top level, e.g.:
    docx.skill  ->  docx/SKILL.md, docx/REFERENCE.md, ...
"""
import argparse
import sys
import zipfile
from pathlib import Path


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("skill_dir")
    ap.add_argument("-o", "--output", default=None,
                    help="output .skill path (default: <skill_dir name>.skill beside it)")
    args = ap.parse_args()

    skill = Path(args.skill_dir).resolve()
    if not skill.is_dir():
        print(f"not a directory: {skill}", file=sys.stderr)
        return 2
    skill_md = skill / "SKILL.md"
    if not skill_md.exists():
        print(f"missing SKILL.md in {skill}", file=sys.stderr)
        return 2

    out = Path(args.output) if args.output else skill.parent / f"{skill.name}.skill"

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(skill.rglob("*")):
            if f.is_file() and not f.name.startswith("."):
                zf.write(f, f.relative_to(skill.parent))

    print(f"packed: {out} ({out.stat().st_size} bytes, {len(list(skill.rglob('*')))} entries)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
