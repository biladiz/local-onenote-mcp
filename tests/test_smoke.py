import sys

import pytest

# simple smoke test verifying the server module loads and exposes expected
# attributes. This does not depend on OneNote being installed.

def test_import_module():
    import onenote_pro_mcp_ps

    assert hasattr(onenote_pro_mcp_ps, "mcp")
    assert onenote_pro_mcp_ps.mcp.name == "local-onenote-mcp"


def test_platform_check(monkeypatch, capsys):
    # simulate non-Windows environment by temporarily changing sys.platform
    monkeypatch.setattr(sys, "platform", "linux")
    # re-import the module in a new namespace to run platform guard
    import importlib
    with pytest.raises(SystemExit) as excinfo:
        importlib.reload(importlib.import_module("onenote_pro_mcp_ps"))
    assert excinfo.value.code == 1
    captured = capsys.readouterr()
    assert "only runs on Windows" in captured.err
