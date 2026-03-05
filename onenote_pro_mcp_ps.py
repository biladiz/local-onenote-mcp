"""
Local OneNote MCP Server (COM-based)
Talks to OneNote exclusively via the PowerShell bridge (onenote_bridge.ps1)
running as a single persistent subprocess in loop mode.

Only runs on Windows (COM automation is Windows-only).
"""

import json
import os
import subprocess
import sys
import threading

# Immediately abort on unsupported platforms to avoid confusing COM errors
if sys.platform != "win32":
    sys.stderr.write("Error: local-onenote-mcp only runs on Windows.\n")
    sys.exit(1)

from fastmcp import FastMCP  # noqa: E402

# ── FastMCP instance ───────────────────────────────────────────────────────

__version__ = "0.1.0"

mcp = FastMCP("local-onenote-mcp")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, "onenote_bridge.ps1")

# ── Persistent PS subprocess manager ──────────────────────────────────────

_process: "subprocess.Popen[str] | None" = None
_lock = threading.Lock()


def _start_process() -> "subprocess.Popen[str]":
    """Spawn the PowerShell bridge in persistent loop mode and wait for its ready signal."""
    proc = subprocess.Popen(
        [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            SCRIPT_PATH,
            "-Loop",
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,
    )
    # Read until we get the {"status":"ready"} handshake
    for _ in range(20):
        line = proc.stdout.readline()
        if not line:
            break
        line = line.strip()
        try:
            resp = json.loads(line)
            if resp.get("status") == "ready":
                return proc
        except json.JSONDecodeError:
            continue  # skip non-JSON startup noise
    raise RuntimeError(
        "PowerShell bridge did not send ready signal. Is OneNote Desktop running?"
    )


def _get_process() -> "subprocess.Popen[str]":
    global _process
    if _process is None or _process.poll() is not None:
        _process = _start_process()
    return _process


def _format_data(data: object) -> str:
    """Convert structured PowerShell output to a human-readable string for MCP."""
    if isinstance(data, list):
        if all(isinstance(x, str) for x in data):
            return "\n".join(data)
        return json.dumps(data, indent=2, ensure_ascii=False)
    if isinstance(data, dict):
        return json.dumps(data, indent=2, ensure_ascii=False)
    return str(data) if data is not None else ""


def run_command(cmd: str, p1: str = "", p2: str = "", p3: str = "") -> str:
    """Send a JSON command to the persistent PS subprocess and return the result."""
    with _lock:
        try:
            proc = _get_process()
            request = json.dumps({"cmd": cmd, "p1": p1, "p2": p2, "p3": p3})
            proc.stdin.write(request + "\n")
            proc.stdin.flush()

            # Read lines until we get a valid JSON response (skipping any warnings)
            for _ in range(30):
                line = proc.stdout.readline()
                if not line:
                    return "Error: PowerShell bridge closed unexpectedly."
                line = line.strip()
                if not line:
                    continue
                try:
                    response = json.loads(line)
                    if response.get("status") == "error":
                        return f"Error: {response.get('error', 'Unknown error')}"
                    return _format_data(response.get("data", ""))
                except json.JSONDecodeError:
                    continue  # skip non-JSON lines (e.g. PS warnings)

            return "Error: No valid response received from PowerShell bridge."

        except Exception as exc:  # noqa: BLE001
            global _process  # noqa: PLW0603
            _process = None  # force restart on next call
            return f"Error: {exc}"


# ── MCP Tools ──────────────────────────────────────────────────────────────


@mcp.tool()
def list_notebooks() -> str:
    """List all LOCAL OneNote notebooks (cloud/OneDrive-only notebooks are excluded)."""
    return run_command("list")


@mcp.tool()
def get_notebook(notebook_name: str) -> str:
    """Get the name and internal ID of a specific local notebook."""
    return run_command("getnotebook", p1=notebook_name)


@mcp.tool()
def list_sections(notebook_name: str) -> str:
    """List all sections in a local notebook."""
    return run_command("sections", p1=notebook_name)


@mcp.tool()
def list_pages(notebook_name: str, section_name: str) -> str:
    """List all pages (name, ID, last-modified) in a specific section of a local notebook."""
    return run_command("listpages", p1=notebook_name, p2=section_name)


@mcp.tool()
def read_notebook(notebook_name: str) -> str:
    """Read the full text content of every page in a local notebook."""
    return run_command("readnotebook", p1=notebook_name)


@mcp.tool()
def read_page(page_id: str) -> str:
    """Read the raw text content of a page given its ID."""
    return run_command("readpage", p1=page_id)


@mcp.tool()
def export_page_as_text(page_id: str) -> str:
    """Export a page as clean plain text — HTML tags and entities stripped. Best for AI consumption."""
    return run_command("exporttext", p1=page_id)


@mcp.tool()
def get_page_metadata(page_id: str) -> str:
    """Return structured metadata for a page: name, section, notebook, and lastModified timestamp."""
    return run_command("pagemetadata", p1=page_id)


@mcp.tool()
def search_pages(query: str) -> str:
    """Search all LOCAL OneNote notebooks for pages matching the query string."""
    return run_command("search", p1=query)


@mcp.tool()
def get_last_updated_page() -> str:
    """Find the single most recently modified page across all local notebooks."""
    return run_command("lastupdated")


@mcp.tool()
def get_last_updated_pages(notebook_name: str = "", limit: int = 5) -> str:
    """Find the N most recently modified items across local notebooks.

    Args:
        notebook_name: Optional. Restrict results to this notebook.
        limit:         Number of items to return (default 5).
    """
    return run_command("lastpages", p1=notebook_name, p2=str(limit))


@mcp.tool()
def create_section(notebook_name: str, section_name: str) -> str:
    """Create a new section in a local notebook."""
    return run_command("createsection", p1=notebook_name, p2=section_name)


@mcp.tool()
def create_page(notebook_name: str, section_name: str, title: str) -> str:
    """Create a new page with the given title in a section of a local notebook."""
    return run_command("createpage", p1=notebook_name, p2=section_name, p3=title)


@mcp.tool()
def update_page(page_id: str, content: str, mode: str = "append") -> str:
    """Write text content to an existing page.

    Args:
        page_id: The ID of the page to update.
        content: Plain text content to write.
        mode:    'append' (default) adds below existing content; 'replace' overwrites it.
    """
    return run_command("updatepage", p1=page_id, p2=content, p3=mode)


# ── Entrypoint ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run(transport="stdio")
