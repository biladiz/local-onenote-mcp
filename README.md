# 📓 Local OneNote MCP Server (COM-based)

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)
![CI](https://github.com/biladiz/local-onenote-mcp/actions/workflows/ci.yml/badge.svg)

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that interacts with your **local Microsoft OneNote Desktop application** via the Windows COM interface. Works exclusively with **local OneNote files** — no cloud permissions or Microsoft Graph API required.

## Features

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all open local OneNote notebooks |
| `get_notebook` | Get name and ID of a specific notebook |
| `list_sections` | List all sections in a notebook |
| `list_pages` | List all pages (with IDs) in a section |
| `read_notebook` | Read all content from a notebook |
| `read_page` | Read text content of a page by ID |
| `export_page_as_text` | Export a page as clean plain text (HTML stripped) |
| `get_page_metadata` | Get name, section, notebook, and last-modified of a page |
| `search_pages` | Search across all local notebooks |
| `get_last_updated_page` | Find the most recently modified page |
| `get_last_updated_pages` | Find the N most recently modified items |
| `create_section` | Create a new section in a notebook |
| `create_page` | Create a new page in a section |
| `update_page` | Append or replace content on an existing page |

## Requirements

- **Windows** (Required — COM automation is Windows-only)
- **Microsoft OneNote Desktop** (2016 / 2019 / 365 Desktop, **not** the UWP Store app)
- **Local Notebooks** (stored on the local filesystem; OneDrive-only notebooks are excluded)
- **Python 3.10+**

## Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/biladiz/local-onenote-mcp
cd local-onenote-mcp
```

### 2. Run the installer

**Windows (Command Prompt / PowerShell):**
```cmd
install.bat
```

**Windows (Git Bash):**
```bash
chmod +x install.sh && ./install.sh
```

The installer:
- ✅ Verifies Python 3.10+
- ✅ Creates a `.venv` virtual environment
- ✅ Installs `fastmcp`
- ✅ Generates `mcp_config_sample.json` with your local paths

### 3. Add to your MCP client

Merge the generated `mcp_config_sample.json` into your client config. For **Gemini Antigravity**:

```
~/.gemini/antigravity/mcp_config.json
```

Add the `"local-onenote-mcp"` entry inside the existing `"mcpServers"` object.

### 4. Restart your MCP client

OneNote Desktop must be **open** before starting the MCP client.

## How It Works

```
MCP Client ──stdio──▸ onenote_pro_mcp_ps.py ──stdin/stdout JSON──▸ onenote_bridge.ps1 [-Loop] ──COM──▸ OneNote
```

1. The MCP server starts the PowerShell bridge **once** at startup (`-Loop` mode).
2. Every tool call sends a JSON command over stdin and reads a JSON response from stdout — no per-call subprocess spawning.
3. The PowerShell bridge uses the **OneNote COM API** and only exposes local notebooks.

## Project Structure

```
local-onenote-mcp/
├── onenote_pro_mcp_ps.py       # MCP server (FastMCP + persistent subprocess manager)
├── onenote_bridge.ps1          # PowerShell bridge (OneNote COM, JSON I/O, loop mode)
├── requirements.txt            # Python dependencies (pinned)
├── install.bat                 # Windows installer
├── install.sh                  # Git Bash installer
├── CONTRIBUTING.md             # Contributor guide
├── .github/workflows/ci.yml    # GitHub Actions (ruff + PSScriptAnalyzer)
├── mcp_config_sample.json      # Generated after install — paste into your MCP config
└── README.md                   # This file
```

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `OneNote.Application` COM error | Open OneNote Desktop before starting the MCP client |
| `Bridge did not send ready signal` | OneNote is not running or COM is blocked — open OneNote first |
| Python not found | Install Python 3.10+ and check "Add to PATH" |
| Permission error on `.ps1` | Installer uses `-ExecutionPolicy Bypass` — run install again |
| Tools return empty results | Ensure notebooks are local files (not OneDrive-only) and open |
| Cloud notebook not visible | Intentional — this server only exposes local/file-based notebooks |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for dev setup, architecture notes, and how to add new tools.

## License

MIT
