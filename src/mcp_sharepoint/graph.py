"""Microsoft Graph client utilities for SharePoint and OneDrive."""

from __future__ import annotations

import base64
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import quote

import requests
from requests import RequestException

from .auth import GraphAuthManager


class GraphAPIError(RuntimeError):
    """Raised when the Microsoft Graph API returns an error response."""

    def __init__(self, message: str, *, status_code: Optional[int] = None, details: Optional[Any] = None) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.details = details


def _is_text_mime(mime_type: Optional[str]) -> bool:
    if not mime_type:
        return False
    if mime_type.startswith("text/"):
        return True
    return mime_type in {
        "application/json",
        "application/xml",
        "application/x-javascript",
        "application/javascript",
        "application/x-httpd-php",
        "application/x-sh",
        "application/x-python",
        "application/sql",
    }


class GraphClient:
    """Thin wrapper around Microsoft Graph REST operations for drive items."""

    def __init__(
        self,
        *,
        auth_manager: GraphAuthManager,
        base_url: Optional[str] = None,
        logger,
    ) -> None:
        self._auth_manager = auth_manager
        self._base_url = base_url or "https://graph.microsoft.com/v1.0"
        self._logger = logger
        self._active_drive_id: Optional[str] = None

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _authorization_headers(self, *, extra: Optional[Dict[str, str]] = None) -> Dict[str, str]:
        token = self._auth_manager.acquire_token_silent()
        headers = {"Authorization": f"Bearer {token}"}
        if extra:
            headers.update(extra)
        return headers

    def _request(
        self,
        method: str,
        endpoint: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Any] = None,
        data: Optional[bytes] = None,
        headers: Optional[Dict[str, str]] = None,
        stream: bool = False,
    ) -> Any:
        url = f"{self._base_url}{endpoint}"
        request_headers = self._authorization_headers(extra=headers)
        if json_body is not None and "Content-Type" not in request_headers and not stream and data is None:
            request_headers["Content-Type"] = "application/json"

        try:
            response = requests.request(
                method,
                url,
                params=params,
                json=json_body,
                data=data,
                headers=request_headers,
                stream=stream,
            )
        except RequestException as exc:  # pragma: no cover - network failure guard
            raise GraphAPIError("Network error while calling Microsoft Graph", details=str(exc)) from exc

        if stream:
            if 200 <= response.status_code < 300:
                return response
            raise GraphAPIError(
                f"Microsoft Graph request failed with status {response.status_code}",
                status_code=response.status_code,
                details=_safe_json(response),
            )

        if 200 <= response.status_code < 300:
            if response.content:
                try:
                    return response.json()
                except ValueError:
                    return response.content
            return {}

        raise GraphAPIError(
            f"Microsoft Graph request failed with status {response.status_code}",
            status_code=response.status_code,
            details=_safe_json(response),
        )

    def _resolve_drive_id(self, drive_id: Optional[str]) -> str:
        drive = drive_id or self._active_drive_id
        if not drive:
            raise GraphAPIError(
                "No drive specified. Use Set_Active_Drive or provide drive_id explicitly in the tool call."
            )
        return drive

    def _resolve_item(self, drive_id: str, path: Optional[str]) -> Dict[str, Any]:
        if not path or path.strip("/") == "":
            return self._request("GET", f"/drives/{drive_id}/root")
        normalized = path.strip("/")
        return self._request("GET", f"/drives/{drive_id}/root:/{normalized}")

    @staticmethod
    def _simplify_drive_item(item: Dict[str, Any]) -> Dict[str, Any]:
        parent = item.get("parentReference", {})
        return {
            "id": item.get("id"),
            "name": item.get("name"),
            "driveId": parent.get("driveId"),
            "path": parent.get("path"),
            "webUrl": item.get("webUrl"),
            "createdDateTime": item.get("createdDateTime"),
            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            "size": item.get("size"),
            "folder": item.get("folder"),
            "file": item.get("file"),
        }

    # ------------------------------------------------------------------
    # Account and drive context helpers
    # ------------------------------------------------------------------
    def set_active_drive(self, drive_id: str) -> Dict[str, Any]:
        metadata = self.get_drive(drive_id)
        self._active_drive_id = drive_id
        return metadata

    def get_active_drive(self) -> Optional[str]:
        return self._active_drive_id

    def get_context(self) -> Dict[str, Any]:
        active_account = self._auth_manager.get_active_account_summary()
        return {
            "active_account": active_account.__dict__ if active_account else None,
            "active_drive_id": self._active_drive_id,
        }

    # ------------------------------------------------------------------
    # Drives and sites
    # ------------------------------------------------------------------
    def list_my_drives(self) -> List[Dict[str, Any]]:
        payload = self._request("GET", "/me/drives")
        return payload.get("value", [])

    def get_drive(self, drive_id: str) -> Dict[str, Any]:
        return self._request("GET", f"/drives/{drive_id}")

    def search_sites(self, query: str) -> List[Dict[str, Any]]:
        payload = self._request("GET", "/sites", params={"search": query})
        return payload.get("value", [])

    def get_site_by_url(self, site_url: str) -> Dict[str, Any]:
        # site_url example: https://tenant.sharepoint.com/sites/Example
        trimmed = site_url.strip()
        if not trimmed:
            raise GraphAPIError("site_url cannot be empty")
        if trimmed.endswith("/"):
            trimmed = trimmed[:-1]
        if "//" not in trimmed:
            raise GraphAPIError("site_url must be an absolute URL")
        host = trimmed.split("//", 1)[1].split("/", 1)[0]
        relative = ""
        if "/" in trimmed.split("//", 1)[1]:
            relative = trimmed.split(host, 1)[1]
        relative = relative.strip("/")
        endpoint = f"/sites/{host}:/{relative}:" if relative else f"/sites/{host}:"
        return self._request("GET", endpoint)

    def list_site_drives(self, *, site_id: Optional[str] = None, site_url: Optional[str] = None) -> List[Dict[str, Any]]:
        if not site_id and not site_url:
            raise GraphAPIError("Either site_id or site_url must be provided to list site drives")
        if site_url and not site_id:
            site_metadata = self.get_site_by_url(site_url)
            site_id = site_metadata.get("id")
            if not site_id:
                raise GraphAPIError("Unable to resolve site ID from the provided URL", details=site_metadata)
        payload = self._request("GET", f"/sites/{site_id}/drives")
        return payload.get("value", [])

    # ------------------------------------------------------------------
    # Drive item operations
    # ------------------------------------------------------------------
    def list_items(self, *, drive_id: Optional[str] = None, path: Optional[str] = None) -> List[Dict[str, Any]]:
        drive = self._resolve_drive_id(drive_id)
        if not path or path.strip("/") == "":
            payload = self._request("GET", f"/drives/{drive}/root/children")
        else:
            normalized = path.strip("/")
            payload = self._request("GET", f"/drives/{drive}/root:/{normalized}:/children")
        items = payload.get("value", [])
        return [self._simplify_drive_item(item) for item in items]

    def get_item_metadata(self, *, drive_id: Optional[str] = None, path: str) -> Dict[str, Any]:
        drive = self._resolve_drive_id(drive_id)
        item = self._resolve_item(drive, path)
        simplified = self._simplify_drive_item(item)
        simplified["raw"] = item
        return simplified

    def get_item_content(self, *, drive_id: Optional[str] = None, path: str) -> Dict[str, Any]:
        drive = self._resolve_drive_id(drive_id)
        item = self._resolve_item(drive, path)
        if "@microsoft.graph.downloadUrl" not in item:
            item = self._request("GET", f"/drives/{drive}/items/{item.get('id')}?select=name,webUrl,size,file,folder,@microsoft.graph.downloadUrl")
        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            raise GraphAPIError("The requested item does not have downloadable content", details=item)

        try:
            response = requests.get(download_url)
        except RequestException as exc:  # pragma: no cover - network failure guard
            raise GraphAPIError("Network error while downloading file content", details=str(exc)) from exc
        if not response.ok:
            raise GraphAPIError(
                f"Failed to download file content (status {response.status_code})",
                status_code=response.status_code,
            )

        mime_type = None
        file_info = item.get("file")
        if file_info:
            mime_type = file_info.get("mimeType")
        else:
            mime_type = response.headers.get("Content-Type")

        content_bytes = response.content
        if _is_text_mime(mime_type):
            try:
                decoded = content_bytes.decode("utf-8")
                content = {"content_type": "text", "content": decoded}
            except UnicodeDecodeError:
                content = {
                    "content_type": "binary",
                    "content_base64": base64.b64encode(content_bytes).decode("ascii"),
                }
        else:
            content = {
                "content_type": "binary",
                "content_base64": base64.b64encode(content_bytes).decode("ascii"),
            }

        return {
            "name": item.get("name"),
            "webUrl": item.get("webUrl"),
            "size": item.get("size"),
            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            **content,
        }

    def create_folder(
        self,
        *,
        folder_name: str,
        parent_path: Optional[str] = None,
        drive_id: Optional[str] = None,
        conflict_behavior: str = "fail",
    ) -> Dict[str, Any]:
        drive = self._resolve_drive_id(drive_id)
        parent = self._resolve_item(drive, parent_path)
        payload = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": conflict_behavior,
        }
        created = self._request(
            "POST",
            f"/drives/{drive}/items/{parent.get('id')}/children",
            json_body=payload,
        )
        return self._simplify_drive_item(created)

    def upload_file(
        self,
        *,
        item_path: str,
        content: str,
        drive_id: Optional[str] = None,
        is_base64: bool = False,
        conflict_behavior: str = "replace",
    ) -> Dict[str, Any]:
        drive = self._resolve_drive_id(drive_id)
        normalized = item_path.strip("/")
        if not normalized:
            raise GraphAPIError("item_path must include the file name")

        file_bytes = base64.b64decode(content) if is_base64 else content.encode("utf-8")
        params = {"@microsoft.graph.conflictBehavior": conflict_behavior}
        result = self._request(
            "PUT",
            f"/drives/{drive}/root:/{normalized}:/content",
            params=params,
            data=file_bytes,
            headers={"Content-Type": "application/octet-stream"},
        )
        return self._simplify_drive_item(result)

    def delete_item(self, *, drive_id: Optional[str] = None, path: str) -> Dict[str, Any]:
        drive = self._resolve_drive_id(drive_id)
        item = self._resolve_item(drive, path)
        self._request("DELETE", f"/drives/{drive}/items/{item.get('id')}")
        return {
            "success": True,
            "id": item.get("id"),
            "name": item.get("name"),
            "path": item.get("parentReference", {}).get("path"),
        }

    # ------------------------------------------------------------------
    # Search operations
    # ------------------------------------------------------------------
    def search_drive_items(
        self,
        *,
        query: str,
        drive_id: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        drive = self._resolve_drive_id(drive_id)
        payload = self._request(
            "GET",
            f"/drives/{drive}/root/search(q='{quote(query)}')",
        )
        items = payload.get("value", [])
        return [self._simplify_drive_item(item) for item in items]

    def search_everywhere(
        self,
        *,
        query: str,
        entity_types: Optional[Iterable[str]] = None,
        size: int = 25,
    ) -> List[Dict[str, Any]]:
        search_body = {
            "requests": [
                {
                    "entityTypes": list(entity_types) if entity_types else ["driveItem", "list", "listItem", "site"],
                    "query": {"queryString": query},
                    "from": 0,
                    "size": size,
                }
            ]
        }
        payload = self._request("POST", "/search/query", json_body=search_body)
        results: List[Dict[str, Any]] = []
        for response in payload.get("value", []):
            hits_container = response.get("hitsContainers", [])
            for container in hits_container:
                for hit in container.get("hits", []):
                    resource = hit.get("resource", {})
                    result_entry = {
                        "name": resource.get("name") or resource.get("title"),
                        "summary": hit.get("summary"),
                        "webUrl": resource.get("webUrl"),
                        "lastModifiedDateTime": resource.get("lastModifiedDateTime"),
                        "size": resource.get("size"),
                        "resourceType": resource.get("@odata.type"),
                    }
                    result_entry.update({k: v for k, v in resource.items() if k not in result_entry})
                    results.append(result_entry)
        return results


def _safe_json(response: requests.Response) -> Any:
    try:
        return response.json()
    except ValueError:
        return response.text

