import subprocess
import json
import os
from fastmcp import FastMCP

mcp = FastMCP("local-onenote-mcp")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, "onenote_bridge.ps1")


def run_ps_command(cmd: str, p1: str = "", p2: str = "", p3: str = "") -> str:
    """Run PowerShell script with parameters and return output."""
    try:
        result = subprocess.run(
            [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                SCRIPT_PATH,
                "-Cmd",
                cmd,
                "-P1",
                p1,
                "-P2",
                p2,
                "-P3",
                p3,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            return f"Error: {result.stderr}"
        return result.stdout.strip()
    except Exception as e:
        return f"Error: {str(e)}"


@mcp.tool()
def list_notebooks() -> str:
    """Returns a list of all open LOCAL OneNote notebook names."""
    return run_ps_command("list")


@mcp.tool()
def get_notebook(notebook_name: str) -> str:
    """Get details about a specific notebook by name."""
    return run_ps_command("getnotebook", p1=notebook_name)


@mcp.tool()
def list_sections(notebook_name: str) -> str:
    """List all sections in a local notebook."""
    return run_ps_command("sections", p1=notebook_name)


@mcp.tool()
def read_notebook(notebook_name: str) -> str:
    """Read all content from a local notebook (all pages)."""
    return run_ps_command("readnotebook", p1=notebook_name)


@mcp.tool()
def search_pages(query: str) -> str:
    """Search for pages across all local notebooks containing the query."""
    return run_ps_command("search", p1=query)


@mcp.tool()
def read_page(page_id: str) -> str:
    """Read content from a specific page by ID."""
    return run_ps_command("readpage", p1=page_id)


@mcp.tool()
def get_last_updated_page() -> str:
    """Find the most recently updated page among local notebooks."""
    return run_ps_command("lastupdated")


@mcp.tool()
def get_last_updated_pages(notebook_name: str = "", limit: int = 5) -> str:
    """Find the last N recently updated local pages, optionally filtered by notebook name."""
    return run_ps_command("lastpages", p1=notebook_name, p2=str(limit))


@mcp.tool()
def create_page(notebook_name: str, section_name: str, title: str) -> str:
    """Create a new page in a section with the given title."""
    return run_ps_command("createpage", p1=notebook_name, p2=section_name, p3=title)


if __name__ == "__main__":
    mcp.run(transport="stdio")
