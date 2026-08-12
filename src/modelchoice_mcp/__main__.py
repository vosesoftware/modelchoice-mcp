"""ModelChoice MCP server entrypoint.

    python -m modelchoice_mcp                              # stdio (default)
    python -m modelchoice_mcp --transport=streamable-http  # HTTP, 127.0.0.1:8000
    python -m modelchoice_mcp --transport=sse --port=9000
"""

from __future__ import annotations

import argparse
from typing import Literal, cast

from modelchoice_mcp.server import mcp

Transport = Literal["stdio", "streamable-http", "sse"]


def main() -> None:
    parser = argparse.ArgumentParser(prog="modelchoice-mcp")
    parser.add_argument(
        "--transport",
        choices=["stdio", "streamable-http", "sse"],
        default="stdio",
        help="MCP transport (default: stdio).",
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    transport = cast(Transport, args.transport)
    if transport == "stdio":
        mcp.run(transport="stdio")
    else:
        # mcp 2.0 takes host/port as run() keyword arguments, forwarded to
        # run_streamable_http_async / run_sse_async. Under 1.x these were set by
        # mutating `mcp.settings` before calling run().
        mcp.run(transport=transport, host=args.host, port=args.port)


if __name__ == "__main__":
    main()
