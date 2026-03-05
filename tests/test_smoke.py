import importlib
import sys

import pytest

# Simple smoke tests verifying the server module loads and exposes expected
# attributes. These do not depend on OneNote being installed.


def test_import_module():
    import onenote_pro_mcp_ps

    assert hasattr(onenote_pro_mcp_ps, "mcp")
    assert onenote_pro_mcp_ps.mcp.name == "local-onenote-mcp"


def test_platform_check(monkeypatch, capsys):
    """Simulate a non-Windows environment and verify the platform guard fires."""
    monkeypatch.setattr(sys, "platform", "linux")
    with pytest.raises(SystemExit) as excinfo:
        importlib.reload(importlib.import_module("onenote_pro_mcp_ps"))
    assert excinfo.value.code == 1
    captured = capsys.readouterr()
    assert "only runs on Windows" in captured.err
