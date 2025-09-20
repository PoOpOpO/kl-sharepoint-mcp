import asyncio

from .common import logger, mcp


async def main() -> None:
    """Entry point for launching the MCP server over stdio."""

    logger.info("Starting SharePoint MCP server ...")

    # Import tools (registration happens at import time)
    from . import tools  # noqa: F401

    logger.info("Running MCP server...")
    await mcp.run_stdio_async()


if __name__ == "__main__":
    asyncio.run(main())

