# Contributing to local-onenote-mcp

Thanks for your interest in contributing! This guide will get you set up quickly.

## Prerequisites

- Windows 10/11
- Microsoft OneNote Desktop (2016 / 2019 / 365) **open and running**
- Python 3.10+
- PowerShell 5.1+

## Dev Setup

```bash
git clone https://github.com/biladiz/local-onenote-mcp
cd local-onenote-mcp
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Project Layout

```
onenote_pro_mcp_ps.py   # FastMCP server — persistent subprocess manager + tool definitions
onenote_bridge.ps1      # PowerShell bridge — all OneNote COM calls, JSON I/O, loop mode
```

## Architecture

```
MCP Client ──stdio──▸ onenote_pro_mcp_ps.py ──stdin/stdout JSON──▸ onenote_bridge.ps1 [-Loop] ──COM──▸ OneNote
```

The Python server starts the PS bridge **once** with `-Loop` at startup. Every tool call sends a single-line JSON command and reads a single-line JSON response — no per-call subprocess spawning.

## Testing Manually

**Test the bridge directly (single-shot mode):**
```powershell
# List local notebooks
powershell -ExecutionPolicy Bypass -File onenote_bridge.ps1 -Cmd list

# List pages in a section
powershell -ExecutionPolicy Bypass -File onenote_bridge.ps1 -Cmd listpages -P1 "MyNotebook" -P2 "MySection"
```

**Test the loop mode:**
```powershell
echo '{"cmd":"list","p1":"","p2":"","p3":""}' | powershell -ExecutionPolicy Bypass -File onenote_bridge.ps1 -Loop
```

## Adding a New Tool

1. **Add a PS command** — add a new `"commandname"` case inside `Invoke-Cmd` in `onenote_bridge.ps1`. Always call `Send-Ok` or `Send-Err`.
2. **Add a Python tool** — add a `@mcp.tool()` decorated function in `onenote_pro_mcp_ps.py` that calls `run_command("commandname", ...)`.
3. **Update the README** — add the new tool to the Features table.

## Code Style

- **Python**: formatted with `ruff format`, linted with `ruff check`. Run: `ruff check . && ruff format --check .`
- **PowerShell**: analysed with `PSScriptAnalyzer`. Run: `Invoke-ScriptAnalyzer -Path onenote_bridge.ps1`

## Submitting a PR

1. Fork the repo and create a feature branch: `git checkout -b feat/my-feature`
2. Make your changes and test manually against OneNote
3. Open a Pull Request with a clear description of what you changed and why

## Reporting Issues

Please include:
- Your OneNote version (File → Account → About OneNote)
- Your Windows version (`winver`)
- The exact error message from the MCP client or from running the bridge manually
