#!/usr/bin/env bash
set -e

if [[ "$(uname)" != "Darwin" ]]; then
  echo "mac-only for now"
  exit 1
fi

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "[1/4] checking uv..."
if ! command -v uv >/dev/null 2>&1; then
  echo "      installing uv..."
  curl -LsSf https://astral.sh/uv/install.sh | sh
  export PATH="$HOME/.local/bin:$PATH"
fi

echo "[2/4] installing ac CLI to ~/.local/bin/ac..."
mkdir -p "$HOME/.local/bin"
install -m 755 "$REPO/bin/ac" "$HOME/.local/bin/ac"

echo "[3/4] installing CLAUDE.md and /archicad command..."
mkdir -p "$HOME/.claude/commands"
python3 "$REPO/scripts/merge_claude_md.py" "$REPO/claude/CLAUDE.md" "$HOME/.claude/CLAUDE.md"
install -m 644 "$REPO/claude/commands/archicad.md" "$HOME/.claude/commands/archicad.md"
# legacy filename — clean up if upgrading from an older install
rm -f "$HOME/.claude/commands/tapir.md"

echo "[4/4] installing ac helper scripts (build_schedule.py)..."
mkdir -p "$HOME/.local/share/ac-scripts"
install -m 644 "$REPO/scripts/build_schedule.py" "$HOME/.local/share/ac-scripts/build_schedule.py"

echo ""
echo "==============================================="
echo "  install complete."
echo ""
echo "  PREREQUISITE: the Tapir Archicad Add-On must be"
echo "  installed inside Archicad — get it from:"
echo "    https://github.com/ENZYME-APD/tapir-archicad-automation"
echo "  Without it, Archicad has no HTTP port for ac to talk to."
echo ""
echo "  Next steps:"
echo "    1. Quit and reopen Claude Code"
echo "    2. Open a project in Archicad"
echo "    3. Run:  ac doctor"
echo "    4. In Claude:  /archicad how many zones"
echo "==============================================="
