# 📓 Local OneNote MCP Server (COM-based)

A Model Context Protocol (MCP) server that interacts with your **local Microsoft OneNote Desktop application** via the COM interface. This server is designed to work exclusively with **local OneNote files** (notebooks stored on your hard drive) and does not require cloud permissions or Microsoft Graph API access.

## Features

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all open local OneNote notebooks |
| `get_notebook` | Get details about a specific notebook |
| `list_sections` | List all sections in a notebook |
| `read_notebook` | Read all content from a notebook |
| `read_page` | Read content from a specific page by ID |
| `search_pages` | Search across all notebooks |
| `get_last_updated_page` | Find the most recently modified page |
| `create_page` | Create a new page in a section |

## Requirements

- **Windows** (Required for COM automation)
- **Microsoft OneNote Desktop App** (The classic version, e.g., OneNote 2016/2019/365 Desktop)
- **Local Notebooks** (Notebooks stored on the local filesystem. OneDrive-only notebooks are excluded)
- **Python 3.10+**

## Quick Start

### 1. Clone or copy this directory

```bash
git clone https://github.com/biladiz/local-onenote-mcp local-onenote-mcp
cd local-onenote-mcp
```

### 2. Run the installer

**Windows (Command Prompt / PowerShell):**
```cmd
install.bat
```

**Windows (Git Bash):**
```bash
chmod +x install.sh
./install.sh
```

The installer will:
- ✅ Verify Python 3.10+ is installed
- ✅ Create a `.venv` virtual environment
- ✅ Install all dependencies (`fastmcp`)
- ✅ Generate a `mcp_config_sample.json` with paths tailored to your system

### 3. Add to your MCP client

Copy the contents of the generated `mcp_config_sample.json` into your MCP client's configuration file. For example, for Gemini Antigravity:

```
~/.gemini/antigravity/mcp_config.json
```

Merge the `"local-onenote-mcp"` entry into the existing `"mcpServers"` object.

### 4. Restart your MCP client

Restart your AI agent / MCP client so it picks up the new server.

## Project Structure

```
local-onenote-mcp/
├── onenote_pro_mcp_ps.py   # MCP server (Python + FastMCP)
├── onenote_bridge.ps1      # PowerShell bridge to OneNote COM API
├── requirements.txt        # Python dependencies
├── install.bat             # Windows installer
├── install.sh              # Bash installer (Git Bash)
├── mcp_config_sample.json  # Generated after install — ready to paste
└── README.md               # This file
```

## How It Works

```
MCP Client ──stdio──▸ onenote_pro_mcp_ps.py ──subprocess──▸ onenote_bridge.ps1 ──COM──▸ OneNote
```

1. The MCP client connects to the Python server via **stdio**.
2. Each tool call runs a PowerShell subprocess with the appropriate command.
3. The PowerShell script uses the **OneNote COM API** to interact with local notebooks only.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `OneNote.Application` COM error | Make sure OneNote desktop is open |
| Python not found | Install Python 3.10+ and add to PATH |
| Permission error on `.ps1` | The installer uses `-ExecutionPolicy Bypass` |
| Tools return empty results | Ensure notebooks are local files (not OneDrive-only) and open |
| Cloud notebook missing | This tool intentionally excludes `https://` based notebooks |

## License

MIT
