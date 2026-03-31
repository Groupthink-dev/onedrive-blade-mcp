"""Microsoft Graph API client for OneDrive & SharePoint operations.

Thin wrapper over httpx — no msgraph-sdk dependency. All methods are synchronous
(the server runs them via asyncio.to_thread for MCP compatibility).
"""

from __future__ import annotations

import base64
import logging
from typing import Any

import httpx

from onedrive_blade_mcp.auth import acquire_token
from onedrive_blade_mcp.models import (
    DEFAULT_LIMIT,
    GRAPH_BASE_URL,
    ITEM_DETAIL_FIELDS,
    ITEM_LIST_FIELDS,
    MAX_UPLOAD_SIZE,
    get_scopes,
    require_write,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Typed exceptions
# ---------------------------------------------------------------------------


class DriveError(Exception):
    """Base exception for Graph API drive errors."""

    def __init__(self, message: str, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class AuthError(DriveError):
    """Authentication or authorisation failure."""


class NotFoundError(DriveError):
    """Resource not found (404)."""


class RateLimitError(DriveError):
    """Graph API throttling (429)."""


class ConflictError(DriveError):
    """Conflict — e.g. name collision on upload/move (409)."""


class WriteDisabledError(DriveError):
    """Write operation attempted while writes are disabled."""


class FileTooLargeError(DriveError):
    """File exceeds simple upload limit."""


# ---------------------------------------------------------------------------
# Error classification
# ---------------------------------------------------------------------------


def _classify_error(resp: httpx.Response) -> DriveError:
    """Map HTTP status to typed exception."""
    try:
        body = resp.json()
        msg = body.get("error", {}).get("message", resp.text[:200])
    except Exception:
        msg = resp.text[:200]

    code = resp.status_code
    if code == 401:
        return AuthError(f"Unauthorized: {msg}", code)
    if code == 403:
        return AuthError(f"Forbidden: {msg}", code)
    if code == 404:
        return NotFoundError(f"Not found: {msg}", code)
    if code == 409:
        return ConflictError(f"Conflict: {msg}", code)
    if code == 429:
        retry = resp.headers.get("Retry-After", "unknown")
        return RateLimitError(f"Rate limited (retry after {retry}s): {msg}", code)
    return DriveError(f"Graph API error {code}: {msg}", code)


# ---------------------------------------------------------------------------
# Token helper
# ---------------------------------------------------------------------------

_scrub_fields = {"access_token", "authorization", "token"}


def _scrub_token(s: str) -> str:
    """Mask bearer tokens in log output."""
    if len(s) > 12:
        return s[:6] + "..." + s[-4:]
    return "***"


# ---------------------------------------------------------------------------
# DriveClient
# ---------------------------------------------------------------------------


class DriveClient:
    """Synchronous Microsoft Graph client for OneDrive/SharePoint."""

    def __init__(self) -> None:
        self._http = httpx.Client(
            base_url=GRAPH_BASE_URL,
            timeout=30.0,
            headers={"Accept": "application/json"},
        )
        self._token: str | None = None

    def close(self) -> None:
        self._http.close()

    # -- Auth ---------------------------------------------------------------

    def _ensure_token(self) -> None:
        if self._token is None:
            self._refresh_token()

    def _refresh_token(self) -> None:
        self._token = acquire_token(get_scopes())
        self._http.headers["Authorization"] = f"Bearer {self._token}"

    # -- HTTP helpers -------------------------------------------------------

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        json_body: dict[str, Any] | None = None,
        content: bytes | None = None,
        content_type: str | None = None,
        expect_json: bool = True,
    ) -> Any:
        self._ensure_token()
        kwargs: dict[str, Any] = {}
        if params:
            kwargs["params"] = params
        if json_body is not None:
            kwargs["json"] = json_body
        if content is not None:
            kwargs["content"] = content
            if content_type:
                kwargs["headers"] = {"Content-Type": content_type}

        resp = self._http.request(method, path, **kwargs)

        # Retry once on 401 (token may have expired)
        if resp.status_code == 401:
            self._refresh_token()
            resp = self._http.request(method, path, **kwargs)

        if resp.status_code >= 400:
            raise _classify_error(resp)

        if expect_json and resp.status_code != 204:
            return resp.json()
        return resp

    def _get(self, path: str, **kwargs: Any) -> Any:
        return self._request("GET", path, **kwargs)

    def _post(self, path: str, **kwargs: Any) -> Any:
        return self._request("POST", path, **kwargs)

    def _patch(self, path: str, **kwargs: Any) -> Any:
        return self._request("PATCH", path, **kwargs)

    def _put(self, path: str, **kwargs: Any) -> Any:
        return self._request("PUT", path, **kwargs)

    def _delete(self, path: str, **kwargs: Any) -> Any:
        return self._request("DELETE", path, expect_json=False, **kwargs)

    # -- Write gate ---------------------------------------------------------

    def _require_write(self) -> None:
        msg = require_write()
        if msg:
            raise WriteDisabledError(msg)

    # ======================================================================
    # READ OPERATIONS
    # ======================================================================

    def get_drive_info(self) -> dict[str, Any]:
        """Get current user's default drive info."""
        return self._get("/me/drive", params={"$select": "id,name,driveType,owner,quota"})

    def list_items(
        self,
        folder_path: str | None = None,
        folder_id: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> dict[str, Any]:
        """List children of a folder.

        Specify either folder_path (e.g. "/Documents/Reports") or folder_id.
        Omit both for root folder.
        """
        select = ",".join(ITEM_LIST_FIELDS)
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 200),
            "$orderby": "name",
        }

        if folder_id:
            path = f"/me/drive/items/{folder_id}/children"
        elif folder_path:
            # Strip leading/trailing slashes for Graph API path format
            clean = folder_path.strip("/")
            path = f"/me/drive/root:/{clean}:/children"
        else:
            path = "/me/drive/root/children"

        return self._get(path, params=params)

    def get_item_metadata(self, item_id: str | None = None, item_path: str | None = None) -> dict[str, Any]:
        """Get detailed metadata for a file or folder."""
        select = ",".join(ITEM_DETAIL_FIELDS)
        params = {"$select": select}

        if item_id:
            path = f"/me/drive/items/{item_id}"
        elif item_path:
            clean = item_path.strip("/")
            path = f"/me/drive/root:/{clean}:"
        else:
            raise DriveError("Either item_id or item_path is required")

        return self._get(path, params=params)

    def read_file_content(
        self,
        item_id: str | None = None,
        item_path: str | None = None,
        max_bytes: int = 1_000_000,
    ) -> dict[str, Any]:
        """Read file content. Returns text for text files, base64 for binary.

        Only reads first max_bytes to prevent memory issues.
        """
        # First get metadata to know the file type and size
        meta = self.get_item_metadata(item_id=item_id, item_path=item_path)
        resolved_id = meta["id"]

        if "folder" in meta:
            raise DriveError("Cannot read content of a folder — use list_items instead")

        size = meta.get("size", 0)
        mime = meta.get("file", {}).get("mimeType", "application/octet-stream")

        # Download content
        self._ensure_token()
        headers = {"Authorization": f"Bearer {self._token}"}
        if size > max_bytes:
            headers["Range"] = f"bytes=0-{max_bytes - 1}"

        resp = self._http.get(
            f"{GRAPH_BASE_URL}/me/drive/items/{resolved_id}/content",
            headers=headers,
            follow_redirects=True,
        )

        if resp.status_code == 401:
            self._refresh_token()
            headers["Authorization"] = f"Bearer {self._token}"
            resp = self._http.get(
                f"{GRAPH_BASE_URL}/me/drive/items/{resolved_id}/content",
                headers=headers,
                follow_redirects=True,
            )

        if resp.status_code >= 400:
            raise _classify_error(resp)

        # Determine if text or binary
        is_text = mime.startswith("text/") or mime in {
            "application/json",
            "application/xml",
            "application/javascript",
            "application/x-yaml",
            "application/x-sh",
        }

        truncated = size > max_bytes

        if is_text:
            try:
                content = resp.text
                return {
                    "metadata": meta,
                    "content_type": "text",
                    "content": content,
                    "truncated": truncated,
                }
            except UnicodeDecodeError:
                pass

        # Binary — return base64
        encoded = base64.b64encode(resp.content).decode("ascii")
        return {
            "metadata": meta,
            "content_type": "binary",
            "content_base64": encoded,
            "truncated": truncated,
            "mime_type": mime,
        }

    def search_files(
        self,
        query: str,
        limit: int = DEFAULT_LIMIT,
    ) -> dict[str, Any]:
        """Search for files across OneDrive."""
        select = ",".join(ITEM_LIST_FIELDS)
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 200),
        }

        return self._get(f"/me/drive/root/search(q='{query}')", params=params)

    def download_url(self, item_id: str | None = None, item_path: str | None = None) -> str:
        """Get a pre-authenticated short-lived download URL for a file."""
        meta = self.get_item_metadata(item_id=item_id, item_path=item_path)
        download = meta.get("@microsoft.graph.downloadUrl")
        if download:
            return download

        # Fall back to createLink if downloadUrl not present
        resolved_id = meta["id"]
        self._ensure_token()
        resp = self._http.get(
            f"{GRAPH_BASE_URL}/me/drive/items/{resolved_id}/content",
            headers={"Authorization": f"Bearer {self._token}"},
            follow_redirects=False,
        )
        if resp.status_code == 302:
            return resp.headers["Location"]
        raise DriveError("Could not obtain download URL")

    def get_versions(self, item_id: str) -> dict[str, Any]:
        """Get version history for a file."""
        return self._get(
            f"/me/drive/items/{item_id}/versions",
            params={"$select": "id,lastModifiedDateTime,lastModifiedBy,size"},
        )

    def get_delta(
        self,
        delta_link: str | None = None,
        folder_path: str | None = None,
    ) -> dict[str, Any]:
        """Track changes using delta queries.

        First call: pass folder_path (or omit for root) to get initial state + delta link.
        Subsequent calls: pass the delta_link from previous response.
        """
        if delta_link:
            # delta_link is a full URL — call directly
            self._ensure_token()
            resp = self._http.get(
                delta_link,
                headers={"Authorization": f"Bearer {self._token}"},
            )
            if resp.status_code == 401:
                self._refresh_token()
                resp = self._http.get(
                    delta_link,
                    headers={"Authorization": f"Bearer {self._token}"},
                )
            if resp.status_code >= 400:
                raise _classify_error(resp)
            return resp.json()

        select = ",".join(ITEM_LIST_FIELDS)
        if folder_path:
            clean = folder_path.strip("/")
            path = f"/me/drive/root:/{clean}:/delta"
        else:
            path = "/me/drive/root/delta"

        return self._get(path, params={"$select": select})

    def get_thumbnail(self, item_id: str, size: str = "medium") -> dict[str, Any]:
        """Get thumbnail URLs for an item.

        size: small (96x96), medium (176x176), large (800x800)
        """
        return self._get(f"/me/drive/items/{item_id}/thumbnails")

    # -- SharePoint ---------------------------------------------------------

    def list_sites(self, query: str | None = None, limit: int = DEFAULT_LIMIT) -> dict[str, Any]:
        """List or search SharePoint sites."""
        params: dict[str, Any] = {
            "$select": "id,name,displayName,webUrl,description",
            "$top": min(limit, 200),
        }

        if query:
            path = f"/sites?search={query}"
            return self._get(path, params=params)
        return self._get("/sites", params=params)

    def list_site_drives(self, site_id: str) -> dict[str, Any]:
        """List document libraries for a SharePoint site."""
        return self._get(
            f"/sites/{site_id}/drives",
            params={"$select": "id,name,driveType,webUrl,quota"},
        )

    def list_site_items(
        self,
        site_id: str,
        drive_id: str,
        folder_id: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> dict[str, Any]:
        """List items in a SharePoint document library folder."""
        select = ",".join(ITEM_LIST_FIELDS)
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 200),
            "$orderby": "name",
        }

        if folder_id:
            path = f"/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children"
        else:
            path = f"/sites/{site_id}/drives/{drive_id}/root/children"

        return self._get(path, params=params)

    # ======================================================================
    # WRITE OPERATIONS
    # ======================================================================

    def upload_file(
        self,
        file_content: bytes,
        file_name: str,
        folder_path: str | None = None,
        folder_id: str | None = None,
        conflict_behavior: str = "rename",
    ) -> dict[str, Any]:
        """Upload a file (simple upload, <=4MB).

        conflict_behavior: rename, replace, or fail
        """
        self._require_write()

        if len(file_content) > MAX_UPLOAD_SIZE:
            raise FileTooLargeError(
                f"File size {len(file_content)} exceeds {MAX_UPLOAD_SIZE} byte limit. "
                "Use resumable upload for larger files."
            )

        if folder_id:
            path = f"/me/drive/items/{folder_id}:/{file_name}:/content"
        elif folder_path:
            clean = folder_path.strip("/")
            path = f"/me/drive/root:/{clean}/{file_name}:/content"
        else:
            path = f"/me/drive/root:/{file_name}:/content"

        params = {"@microsoft.graph.conflictBehavior": conflict_behavior}

        return self._put(
            path,
            content=file_content,
            content_type="application/octet-stream",
            params=params,
        )

    def create_folder(
        self,
        name: str,
        parent_path: str | None = None,
        parent_id: str | None = None,
    ) -> dict[str, Any]:
        """Create a new folder."""
        self._require_write()

        body: dict[str, Any] = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }

        if parent_id:
            path = f"/me/drive/items/{parent_id}/children"
        elif parent_path:
            clean = parent_path.strip("/")
            path = f"/me/drive/root:/{clean}:/children"
        else:
            path = "/me/drive/root/children"

        return self._post(path, json_body=body)

    def move_item(
        self,
        item_id: str,
        destination_folder_id: str | None = None,
        new_name: str | None = None,
    ) -> dict[str, Any]:
        """Move and/or rename a file or folder."""
        self._require_write()

        body: dict[str, Any] = {}
        if destination_folder_id:
            body["parentReference"] = {"id": destination_folder_id}
        if new_name:
            body["name"] = new_name

        if not body:
            raise DriveError("Either destination_folder_id or new_name is required")

        return self._patch(f"/me/drive/items/{item_id}", json_body=body)

    def delete_item(self, item_id: str) -> None:
        """Delete a file or folder (moves to recycle bin)."""
        self._require_write()
        self._delete(f"/me/drive/items/{item_id}")

    def create_sharing_link(
        self,
        item_id: str,
        link_type: str = "view",
        scope: str = "anonymous",
        expiration: str | None = None,
    ) -> dict[str, Any]:
        """Create a sharing link for a file or folder.

        link_type: view, edit, embed
        scope: anonymous, organization
        """
        self._require_write()

        body: dict[str, Any] = {
            "type": link_type,
            "scope": scope,
        }
        if expiration:
            body["expirationDateTime"] = expiration

        return self._post(f"/me/drive/items/{item_id}/createLink", json_body=body)

    def get_permissions(self, item_id: str) -> dict[str, Any]:
        """Get sharing permissions for a file or folder."""
        return self._get(f"/me/drive/items/{item_id}/permissions")


# ---------------------------------------------------------------------------
# Singleton
# ---------------------------------------------------------------------------

_client: DriveClient | None = None


def get_client() -> DriveClient:
    global _client  # noqa: PLW0603
    if _client is None:
        _client = DriveClient()
    return _client
