"""Server boot + tool-registration smoke tests."""

from __future__ import annotations

from modelchoice_mcp import __version__
from modelchoice_mcp.server import mcp


def test_version_set() -> None:
    assert __version__ == "0.0.12"


def test_server_name() -> None:
    assert mcp.name == "modelchoice-mcp"


async def test_read_tools_registered() -> None:
    names = {t.name for t in await mcp.list_tools()}
    assert {"list_trees", "get_tree", "roll_up"} <= names


async def test_tool_descriptions_have_brand_prefix() -> None:
    for t in await mcp.list_tools():
        assert (t.description or "").startswith("ModelChoice:")


async def test_design_prompt_registered() -> None:
    names = {p.name for p in await mcp.list_prompts()}
    assert "design-decision-tree" in names
