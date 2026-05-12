#!/usr/bin/env bash
set -e

echo "stopping ac daemon if running..."
"$HOME/.local/bin/ac" stop 2>/dev/null || true

echo "removing ac CLI..."
rm -f "$HOME/.local/bin/ac"

echo "removing /archicad slash command..."
rm -f "$HOME/.claude/commands/archicad.md"

echo "removing ac helper scripts..."
rm -rf "$HOME/.local/share/ac-scripts"

echo "removing archicad section from CLAUDE.md..."
python3 - <<'PY'
from pathlib import Path
START = "<!-- ARCHICAD-FAST-START -->"
END = "<!-- ARCHICAD-FAST-END -->"
p = Path.home() / ".claude/CLAUDE.md"
if p.exists():
    text = p.read_text()
    if START in text and END in text:
        before = text.split(START)[0].rstrip()
        after = text.split(END, 1)[1].lstrip()
        new = (before + "\n\n" + after).strip()
        if new:
            p.write_text(new + "\n")
        else:
            p.unlink()
PY

echo ""
echo "done."
