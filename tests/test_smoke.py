import subprocess
import sys

# Simple smoke tests verifying the server module loads and exposes expected
# attributes. These do not depend on OneNote being installed.


def test_import_module():
    """Server module must import cleanly on Windows."""
    import onenote_pro_mcp_ps

    assert hasattr(onenote_pro_mcp_ps, "mcp")
    assert onenote_pro_mcp_ps.mcp.name == "local-onenote-mcp"


def test_platform_check():
    """Platform guard must exit with code 1 on non-Windows."""
    # Run in a fresh subprocess so we don't corrupt the current interpreter's
    # module cache or trigger SystemExit inside pytest's own process.
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            "import sys; sys.platform = 'linux'; import onenote_pro_mcp_ps",
        ],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 1
    assert "only runs on Windows" in result.stderr
