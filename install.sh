#!/usr/bin/env bash
set -euo pipefail

TARGET="${1:-antigravity}"
WORKSPACE_DIR="${2:-.}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

install_to() {
  local dest="$1"
  rm -rf "$dest"
  mkdir -p "$dest"
  for item in SKILL.md LICENSE NOTICE requirements.txt scripts references examples; do
    if [ -e "$SCRIPT_DIR/$item" ]; then
      cp -R "$SCRIPT_DIR/$item" "$dest/$item"
    fi
  done
  find "$dest" -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
  echo "Installed morphdeck-3d -> $dest"
}

case "$TARGET" in
  antigravity)
    install_to "$HOME/.gemini/config/skills/morphdeck-3d"
    ;;
  claude)
    install_to "$HOME/.claude/skills/morphdeck-3d"
    ;;
  codex)
    install_to "$HOME/.codex/skills/morphdeck-3d"
    ;;
  cursor)
    install_to "$WORKSPACE_DIR/.cursor/skills/morphdeck-3d"
    ;;
  all)
    install_to "$HOME/.gemini/config/skills/morphdeck-3d"
    install_to "$HOME/.claude/skills/morphdeck-3d"
    install_to "$HOME/.codex/skills/morphdeck-3d"
    install_to "$WORKSPACE_DIR/.cursor/skills/morphdeck-3d"
    ;;
  *)
    echo "Usage: ./install.sh [antigravity|claude|codex|cursor|all] [workspace_dir]" >&2
    exit 1
    ;;
esac
