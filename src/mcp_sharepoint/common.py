"""Common configuration and shared singletons for the SharePoint/OneDrive MCP server."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

from .auth import GraphAuthManager
from .graph import GraphClient

# Load environment variables early so the whole package sees them
load_dotenv()

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
LOG_LEVEL = os.getenv("MCP_GRAPH_LOG_LEVEL", "INFO").upper()
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = os.getenv("MCP_GRAPH_LOG_FILE", "mcp_sharepoint.log")

logging.basicConfig(
    level=LOG_LEVEL,
    format=LOG_FORMAT,
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()],
)

logger = logging.getLogger("mcp_sharepoint")

# ---------------------------------------------------------------------------
# MCP server bootstrap
# ---------------------------------------------------------------------------
SERVER_NAME = os.getenv("MCP_SERVER_NAME", "mcp_sharepoint")
SERVER_INSTRUCTIONS = os.getenv(
    "MCP_SERVER_INSTRUCTIONS",
    (
        "Interact with Microsoft 365 content (SharePoint Online and OneDrive) "
        "using the Microsoft account that is signed into this device. Use the "
        "authentication tools to choose which account to operate with and then "
        "browse, search, create or update files across drives."
    ),
)

mcp = FastMCP(name=SERVER_NAME, instructions=SERVER_INSTRUCTIONS)

# ---------------------------------------------------------------------------
# Authentication and Microsoft Graph client singletons
# ---------------------------------------------------------------------------
DEFAULT_SCOPES = [
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read",
    "offline_access",
]

scopes_env = os.getenv("MCP_GRAPH_SCOPES")
if scopes_env:
    scopes_list = [scope.strip() for scope in scopes_env.split(",") if scope.strip()]
    scopes = scopes_list or DEFAULT_SCOPES
else:
    scopes = DEFAULT_SCOPES

client_id = os.getenv("MCP_GRAPH_CLIENT_ID") or os.getenv("SHP_ID_APP")
if not client_id:
    raise ValueError(
        "The MCP Graph client requires the environment variable MCP_GRAPH_CLIENT_ID "
        "(or legacy SHP_ID_APP) to be set to an Azure AD public client ID."
    )

tenant_id = os.getenv("MCP_GRAPH_TENANT_ID") or os.getenv("SHP_TENANT_ID") or "common"
cache_path = os.getenv("MCP_GRAPH_CACHE_PATH") or os.path.join(
    Path.home(), ".cache", "mcp_sharepoint", "token_cache.bin"
)
os.makedirs(os.path.dirname(cache_path), exist_ok=True)

auth_manager = GraphAuthManager(
    client_id=client_id,
    tenant_id=tenant_id,
    scopes=scopes,
    cache_path=cache_path,
    logger=logger,
)

graph_client = GraphClient(
    auth_manager=auth_manager,
    base_url=os.getenv("MCP_GRAPH_BASE_URL"),
    logger=logger,
)

# Allow pre-selection of a drive through configuration (optional)
preselected_drive = os.getenv("MCP_GRAPH_DEFAULT_DRIVE_ID")
if preselected_drive:
    try:
        graph_client.set_active_drive(preselected_drive)
    except Exception as exc:  # pragma: no cover - defensive configuration guard
        logger.warning("Failed to pre-select drive %s: %s", preselected_drive, exc)

