"""FastMCP server entrypoint.

Constructs the ``mcp`` instance and triggers tool registration by
importing ``modelchoice_mcp.tools`` (the tools attach via the
``@mcp.tool`` decorator side-effect). Read-only in Phase 0/1.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from modelchoice_mcp import __version__

mcp = FastMCP(
    name="modelchoice-mcp",
    instructions=(
        "ModelChoice MCP exposes Vose Software's ModelChoice decision-tree "
        "add-in for Excel. It reads decision trees stored in a workbook and "
        "rolls them back to expected values and the optimal policy — the "
        "decision recommendation — without needing the add-in loaded. Use "
        "list_trees to discover trees, get_tree for structure, and roll_up "
        "for the recommendation."
    ),
)

# Side-effect import registers every @mcp.tool. Must come after `mcp`.
from modelchoice_mcp import tools  # noqa: E402, F401

__all__ = ["__version__", "mcp"]
