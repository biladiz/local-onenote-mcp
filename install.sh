#!/usr/bin/env bash
set -euo pipefail

echo "============================================"
echo "  Local OneNote MCP Server - Installation"
echo "============================================"
echo ""

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${SCRIPT_DIR}/.venv"

# 1. Check Python
echo "[1/4] Checking Python..."
PYTHON_CMD=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then PYTHON_CMD="$cmd"; break; fi
done
[ -z "$PYTHON_CMD" ] && { echo "[ERROR] Python not found."; exit 1; }
PYVER=$("$PYTHON_CMD" --version 2>&1 | awk '{print $2}')
PYMAJOR=$(echo "$PYVER" | cut -d. -f1)
PYMINOR=$(echo "$PYVER" | cut -d. -f2)
echo "       Found $PYTHON_CMD $PYVER"
{ [ "$PYMAJOR" -lt 3 ] || { [ "$PYMAJOR" -eq 3 ] && [ "$PYMINOR" -lt 10 ]; }; } && { echo "[ERROR] Python 3.10+ required."; exit 1; }
echo "       [OK] Version compatible."
echo ""

# 2. OS check
echo "[1.5] Checking OS..."
OS_TYPE="$(uname -s 2>/dev/null || echo Unknown)"
case "$OS_TYPE" in
    MINGW*|MSYS*|CYGWIN*|Windows_NT) echo "       [OK] Windows detected."; PLATFORM="windows" ;;
    *) echo "       [WARNING] Non-Windows: COM automation won't work."; PLATFORM="other" ;;
esac
echo ""

# 3. Create venv
echo "[2/4] Creating virtual environment..."
if [ -f "${VENV_DIR}/bin/python" ] || [ -f "${VENV_DIR}/Scripts/python.exe" ]; then
    echo "       Already exists, skipping."
else
    $PYTHON_CMD -m venv "$VENV_DIR"
    echo "       [OK] Created at ${VENV_DIR}"
fi
echo ""

# 4. Install deps
echo "[3/4] Installing dependencies..."
[ "$PLATFORM" = "windows" ] && PIP_CMD="${VENV_DIR}/Scripts/pip.exe" VENV_PYTHON="${VENV_DIR}/Scripts/python.exe" || PIP_CMD="${VENV_DIR}/bin/pip" VENV_PYTHON="${VENV_DIR}/bin/python"
$PIP_CMD install --upgrade pip >/dev/null 2>&1
$PIP_CMD install -r "${SCRIPT_DIR}/requirements.txt"
echo "       [OK] Done."
echo ""

# 5. Generate MCP config
echo "[4/4] Generating MCP configuration..."
MCP_SCRIPT="${SCRIPT_DIR}/onenote_pro_mcp_ps.py"
[ "$PLATFORM" = "windows" ] && JSON_PYTHON=$(echo "$VENV_PYTHON" | sed 's/\\/\\\\/g') JSON_SCRIPT=$(echo "$MCP_SCRIPT" | sed 's/\\/\\\\/g') || JSON_PYTHON="$VENV_PYTHON" JSON_SCRIPT="$MCP_SCRIPT"
cat > "${SCRIPT_DIR}/mcp_config_sample.json" <<EOF
{
  "mcpServers": {
    "local-onenote-mcp": {
      "command": "${JSON_PYTHON}",
      "args": ["${JSON_SCRIPT}"],
      "disabled": false
    }
  }
}
EOF
echo "       [OK] Written to mcp_config_sample.json"
echo ""

echo "============================================"
echo "  Done! Paste mcp_config_sample.json into"
echo "  your MCP client config, then restart."
echo "============================================"
