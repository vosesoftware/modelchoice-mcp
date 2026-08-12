"""MCP server entrypoint.

Constructs the ``mcp`` instance and triggers tool registration by
importing ``modelchoice_mcp.tools`` (the tools attach via the
``@mcp.tool`` decorator side-effect).

The high-level server class was called ``FastMCP`` and lived in
``mcp.server.fastmcp`` up to mcp 1.x; mcp 2.0 removed that module and renamed the
class to :class:`~mcp.server.MCPServer`. The decorator API is unchanged, so the
rename is the whole migration for this module.
"""

from __future__ import annotations

from mcp.server import MCPServer

from modelchoice_mcp import __version__

mcp = MCPServer(
    name="modelchoice-mcp",
    version=__version__,
    instructions=(
        "ModelChoice MCP exposes Vose Software's ModelChoice decision-tree "
        "add-in for Excel. It reads decision trees stored in a workbook and "
        "rolls them back to expected values and the optimal policy — the "
        "decision recommendation — without needing the add-in loaded. Use "
        "list_trees to discover trees, get_tree for structure, and roll_up "
        "for the recommendation."
    ),
)

# Side-effect imports register every @mcp.tool / @mcp.prompt. Must come
# after `mcp` is constructed.
from modelchoice_mcp import prompts, resources, tools  # noqa: E402, F401

__all__ = ["__version__", "mcp"]
