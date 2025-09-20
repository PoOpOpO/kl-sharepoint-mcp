"""MCP tool definitions for interacting with Microsoft 365 resources."""

from __future__ import annotations

import asyncio
from typing import Iterable, List, Optional

from .auth import AuthenticationError, AuthenticationFlowNotFound
from .common import auth_manager, graph_client, logger, mcp
from .graph import GraphAPIError


async def _run_async(func, *args, **kwargs):
    return await asyncio.to_thread(func, *args, **kwargs)


def _error_response(exc: Exception, *, operation: str) -> dict:
    logger.error("%s failed: %s", operation, exc)
    return {
        "success": False,
        "operation": operation,
        "error": exc.__class__.__name__,
        "message": str(exc),
    }


# ---------------------------------------------------------------------------
# Authentication tools
# ---------------------------------------------------------------------------
@mcp.tool(
    name="Start_Device_Login",
    description=(
        "Inicia un flujo de autenticación con código de dispositivo para elegir "
        "la cuenta de Microsoft 365 que se utilizará en esta sesión."
    ),
)
async def start_device_login():
    try:
        return await _run_async(auth_manager.start_device_login)
    except AuthenticationError as exc:
        return _error_response(exc, operation="start_device_login")


@mcp.tool(
    name="Complete_Device_Login",
    description=(
        "Completa el flujo de autenticación iniciado previamente con Start_Device_Login."),
)
async def complete_device_login(flow_id: str, timeout_seconds: Optional[int] = None):
    try:
        return await _run_async(auth_manager.complete_device_login, flow_id, timeout=timeout_seconds)
    except AuthenticationFlowNotFound as exc:
        return _error_response(exc, operation="complete_device_login")
    except AuthenticationError as exc:
        return _error_response(exc, operation="complete_device_login")


@mcp.tool(
    name="List_Available_Accounts",
    description="Lista las cuentas de Microsoft disponibles en la caché local de este dispositivo.",
)
async def list_available_accounts():
    accounts = await _run_async(auth_manager.list_accounts)
    return [account.__dict__ for account in accounts]


@mcp.tool(
    name="Set_Active_Account",
    description="Selecciona la cuenta de Microsoft 365 con la que operará el conector.",
)
async def set_active_account(
    home_account_id: Optional[str] = None,
    username: Optional[str] = None,
):
    try:
        summary = await _run_async(auth_manager.set_active_account, home_account_id=home_account_id, username=username)
        return {
            "success": True,
            "account": summary.__dict__ if summary else None,
        }
    except AuthenticationError as exc:
        return _error_response(exc, operation="set_active_account")


@mcp.tool(
    name="Get_Auth_Context",
    description="Devuelve información de depuración sobre el estado de autenticación actual.",
)
async def get_auth_context():
    return await _run_async(auth_manager.get_context)


# ---------------------------------------------------------------------------
# Drive and site discovery tools
# ---------------------------------------------------------------------------
@mcp.tool(
    name="List_My_Drives",
    description="Lista todos los drives (OneDrive personal y bibliotecas de documentos) disponibles para la cuenta activa.",
)
async def list_my_drives():
    try:
        return await _run_async(graph_client.list_my_drives)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="list_my_drives")


@mcp.tool(
    name="Search_SharePoint_Sites",
    description="Busca sitios de SharePoint accesibles para la cuenta activa.",
)
async def search_sharepoint_sites(query: str):
    try:
        return await _run_async(graph_client.search_sites, query)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="search_sharepoint_sites")


@mcp.tool(
    name="List_Site_Drives",
    description="Obtiene las bibliotecas de documentos (drives) disponibles en un sitio de SharePoint específico.",
)
async def list_site_drives(site_id: Optional[str] = None, site_url: Optional[str] = None):
    try:
        return await _run_async(graph_client.list_site_drives, site_id=site_id, site_url=site_url)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="list_site_drives")


@mcp.tool(
    name="Set_Active_Drive",
    description="Selecciona el drive sobre el que se ejecutarán las operaciones de archivos.",
)
async def set_active_drive(drive_id: str):
    try:
        metadata = await _run_async(graph_client.set_active_drive, drive_id)
        return {"success": True, "drive": metadata}
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="set_active_drive")


@mcp.tool(
    name="Get_Graph_Context",
    description="Muestra el contexto activo (cuenta y drive seleccionado) para Microsoft Graph.",
)
async def get_graph_context():
    return await _run_async(graph_client.get_context)


# ---------------------------------------------------------------------------
# Drive item operations
# ---------------------------------------------------------------------------
@mcp.tool(
    name="List_Drive_Items",
    description="Lista los elementos dentro de una carpeta del drive activo o de un drive especificado.",
)
async def list_drive_items(path: Optional[str] = None, drive_id: Optional[str] = None):
    try:
        return await _run_async(graph_client.list_items, drive_id=drive_id, path=path)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="list_drive_items")


@mcp.tool(
    name="Get_Drive_Item_Metadata",
    description="Obtiene metadatos detallados de un elemento del drive.",
)
async def get_drive_item_metadata(path: str, drive_id: Optional[str] = None):
    try:
        return await _run_async(graph_client.get_item_metadata, drive_id=drive_id, path=path)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="get_drive_item_metadata")


@mcp.tool(
    name="Get_Drive_Item_Content",
    description="Recupera el contenido de un archivo del drive (en texto plano o base64 según corresponda).",
)
async def get_drive_item_content(path: str, drive_id: Optional[str] = None):
    try:
        return await _run_async(graph_client.get_item_content, drive_id=drive_id, path=path)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="get_drive_item_content")


@mcp.tool(
    name="Create_Drive_Folder",
    description="Crea una carpeta nueva en el drive activo o en el drive indicado.",
)
async def create_drive_folder(
    folder_name: str,
    parent_path: Optional[str] = None,
    drive_id: Optional[str] = None,
    conflict_behavior: str = "fail",
):
    try:
        return await _run_async(
            graph_client.create_folder,
            folder_name=folder_name,
            parent_path=parent_path,
            drive_id=drive_id,
            conflict_behavior=conflict_behavior,
        )
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="create_drive_folder")


@mcp.tool(
    name="Upload_Drive_File",
    description="Carga un archivo nuevo en el drive (permite contenido en texto o en base64).",
)
async def upload_drive_file(
    item_path: str,
    content: str,
    drive_id: Optional[str] = None,
    is_base64: bool = False,
    conflict_behavior: str = "fail",
):
    try:
        return await _run_async(
            graph_client.upload_file,
            item_path=item_path,
            content=content,
            drive_id=drive_id,
            is_base64=is_base64,
            conflict_behavior=conflict_behavior,
        )
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="upload_drive_file")


@mcp.tool(
    name="Update_Drive_File",
    description="Actualiza el contenido de un archivo existente en el drive.",
)
async def update_drive_file(
    item_path: str,
    content: str,
    drive_id: Optional[str] = None,
    is_base64: bool = False,
):
    try:
        return await _run_async(
            graph_client.upload_file,
            item_path=item_path,
            content=content,
            drive_id=drive_id,
            is_base64=is_base64,
            conflict_behavior="replace",
        )
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="update_drive_file")


@mcp.tool(
    name="Delete_Drive_Item",
    description="Elimina un archivo o carpeta del drive indicado.",
)
async def delete_drive_item(path: str, drive_id: Optional[str] = None):
    try:
        return await _run_async(graph_client.delete_item, drive_id=drive_id, path=path)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="delete_drive_item")


@mcp.tool(
    name="Search_Drive_Items",
    description="Busca archivos y carpetas dentro del drive seleccionado.",
)
async def search_drive_items(query: str, drive_id: Optional[str] = None):
    try:
        return await _run_async(graph_client.search_drive_items, query=query, drive_id=drive_id)
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="search_drive_items")


@mcp.tool(
    name="Deep_Search_Microsoft365",
    description=(
        "Realiza una búsqueda global en SharePoint y OneDrive utilizando Microsoft Graph Search "
        "para realizar investigaciones profundas sobre los recursos disponibles."
    ),
)
async def deep_search_microsoft365(
    query: str,
    entity_types: Optional[List[str]] = None,
    size: int = 25,
):
    try:
        entities: Optional[Iterable[str]] = entity_types
        return await _run_async(
            graph_client.search_everywhere,
            query=query,
            entity_types=entities,
            size=size,
        )
    except (AuthenticationError, GraphAPIError) as exc:
        return _error_response(exc, operation="deep_search_microsoft365")

