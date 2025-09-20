"""Package entry point for the Microsoft 365 MCP server."""

import asyncio

from . import server


def main() -> None:
    """Console script entry point."""

    asyncio.run(server.main())


__all__ = ["main", "server"]

