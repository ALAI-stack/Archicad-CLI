#!/usr/bin/env python3
"""Insert or replace the ARCHICAD-FAST section in ~/.claude/CLAUDE.md.

Safe to re-run. Preserves any other content the user has in CLAUDE.md.
"""
import sys
from pathlib import Path

START = "<!-- ARCHICAD-FAST-START -->"
END = "<!-- ARCHICAD-FAST-END -->"

src = Path(sys.argv[1]).read_text().strip()
dst = Path(sys.argv[2])

if dst.exists():
    existing = dst.read_text()
    if START in existing and END in existing:
        before = existing.split(START)[0].rstrip()
        after = existing.split(END, 1)[1].lstrip()
        new = (before + "\n\n" + src + "\n\n" + after).strip() + "\n"
    else:
        new = existing.rstrip() + "\n\n" + src + "\n"
else:
    new = src + "\n"

dst.parent.mkdir(parents=True, exist_ok=True)
dst.write_text(new)
print(f"      wrote {dst}")
