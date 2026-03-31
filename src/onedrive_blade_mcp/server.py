"""OneDrive & SharePoint Blade MCP Server.

16 tools for file operations via Microsoft Graph API.
Follows the Sidereal blade MCP pattern: FastMCP 2.0, lazy-init client,
async wrapper over sync httpx, bearer auth middleware for HTTP transport.
"""

from __future__ import annotations

import asyncio
import base64
import logging
import os

from fastmcp import FastMCP

from onedrive_blade_mcp.auth import BearerAuthMiddleware
from onedrive_blade_mcp.client import DriveError, WriteDisabledError, get_client
from onedrive_blade_mcp.formatters import (
    format_delta,
    format_drive_info,
    format_file_content,
    format_item_detail,
    format_item_list,
    format_permissions,
    format_search_results,
    format_sharing_link,
    format_site_drives,
    format_sites,
    format_versions,
)
from onedrive_blade_mcp.models import DEFAULT_LIMIT

logger = logging.getLogger(__name__)

mcp = FastMCP("OneDrive Blade MCP")

# ---------------------------------------------------------------------------
# Async wrapper
# ---------------------------------------------------------------------------


async def _run(fn, *args, **kwargs) -> str:
    """Run a sync client method in a thread, return formatted result or error."""
    try:
        result = await asyncio.to_thread(fn, *args, **kwargs)
        return result
    except WriteDisabledError as exc:
        return str(exc)
    except DriveError as exc:
        return f"Error: {exc}"
    except Exception as exc:
        logger.exception("Unexpected error in %s", fn.__name__)
        return f"Error: {type(exc).__name__}: {exc}"


# ======================================================================
# META
# ======================================================================


@mcp.tool()
async def drive_info() -> str:
    """Get OneDrive storage info — drive type, quota usage, owner.

    Quick status check; no parameters needed.
    """

    def _do():
        client = get_client()
        data = client.get_drive_info()
        return format_drive_info(data)

    return await _run(_do)


# ======================================================================
# READ TOOLS
# ======================================================================


@mcp.tool()
async def drive_list(
    folder_path: str | None = None,
    folder_id: str | None = None,
    limit: int = DEFAULT_LIMIT,
) -> str:
    """List files and folders in a OneDrive directory.

    Args:
        folder_path: Path like "/Documents/Reports" (omit for root)
        folder_id: Item ID of the folder (alternative to path)
        limit: Max items to return (default 20, max 200)
    """

    def _do():
        client = get_client()
        data = client.list_items(folder_path=folder_path, folder_id=folder_id, limit=limit)
        label = folder_path or folder_id or "Root"
        return format_item_list(data, folder_label=label)

    return await _run(_do)


@mcp.tool()
async def drive_read(
    item_id: str | None = None,
    item_path: str | None = None,
    max_bytes: int = 1_000_000,
) -> str:
    """Read file content from OneDrive.

    Returns text for text files, base64 summary for binary. Reads up to
    max_bytes to prevent memory issues.

    Args:
        item_id: File item ID
        item_path: File path like "/Documents/notes.txt"
        max_bytes: Maximum bytes to read (default 1MB)
    """

    def _do():
        client = get_client()
        data = client.read_file_content(item_id=item_id, item_path=item_path, max_bytes=max_bytes)
        return format_file_content(data)

    return await _run(_do)


@mcp.tool()
async def drive_metadata(
    item_id: str | None = None,
    item_path: str | None = None,
) -> str:
    """Get detailed metadata for a file or folder.

    Returns: name, size, dates, creator, path, sharing status, MIME type.

    Args:
        item_id: Item ID
        item_path: Path like "/Documents/report.pdf"
    """

    def _do():
        client = get_client()
        data = client.get_item_metadata(item_id=item_id, item_path=item_path)
        return format_item_detail(data)

    return await _run(_do)


@mcp.tool()
async def drive_search(query: str, limit: int = DEFAULT_LIMIT) -> str:
    """Search for files across OneDrive by name or content.

    Uses Microsoft Search — matches file names, content, and metadata.

    Args:
        query: Search query (e.g. "budget 2026", "*.pdf")
        limit: Max results (default 20)
    """

    def _do():
        client = get_client()
        data = client.search_files(query=query, limit=limit)
        return format_search_results(data, query)

    return await _run(_do)


@mcp.tool()
async def drive_download(
    item_id: str | None = None,
    item_path: str | None = None,
) -> str:
    """Get a pre-authenticated download URL for a file.

    Returns a short-lived URL that can be used to download the file directly.

    Args:
        item_id: File item ID
        item_path: File path like "/Documents/report.pdf"
    """

    def _do():
        client = get_client()
        url = client.download_url(item_id=item_id, item_path=item_path)
        return f"Download URL (short-lived):\n{url}"

    return await _run(_do)


@mcp.tool()
async def drive_versions(item_id: str) -> str:
    """Get version history for a file.

    Shows all versions with dates, authors, and sizes.

    Args:
        item_id: File item ID
    """

    def _do():
        client = get_client()
        data = client.get_versions(item_id)
        return format_versions(data)

    return await _run(_do)


@mcp.tool()
async def drive_delta(
    delta_link: str | None = None,
    folder_path: str | None = None,
) -> str:
    """Track file changes using delta queries (incremental sync).

    First call: pass folder_path (or omit for root) — returns current state + delta_link.
    Subsequent calls: pass the delta_link — returns only changes since last call.

    Args:
        delta_link: Delta link from previous call (for incremental updates)
        folder_path: Folder to track (omit for root, ignored if delta_link provided)
    """

    def _do():
        client = get_client()
        data = client.get_delta(delta_link=delta_link, folder_path=folder_path)
        return format_delta(data)

    return await _run(_do)


@mcp.tool()
async def drive_thumbnail(item_id: str) -> str:
    """Get thumbnail URLs for a file (images, documents, videos).

    Returns URLs for small (96x96), medium (176x176), and large (800x800) thumbnails.

    Args:
        item_id: File item ID
    """

    def _do():
        client = get_client()
        data = client.get_thumbnail(item_id)
        sets = data.get("value", [])
        if not sets:
            return "No thumbnails available for this item."

        lines = ["Thumbnails:"]
        for ts in sets:
            for size_name in ("small", "medium", "large"):
                thumb = ts.get(size_name)
                if thumb:
                    w = thumb.get("width", "?")
                    h = thumb.get("height", "?")
                    url = thumb.get("url", "")
                    lines.append(f"  {size_name} ({w}x{h}): {url}")
        return "\n".join(lines)

    return await _run(_do)


# -- SharePoint -------------------------------------------------------------


@mcp.tool()
async def drive_sites(query: str | None = None, limit: int = DEFAULT_LIMIT) -> str:
    """List or search SharePoint sites.

    Args:
        query: Optional search query to filter sites
        limit: Max results (default 20)
    """

    def _do():
        client = get_client()
        data = client.list_sites(query=query, limit=limit)
        return format_sites(data)

    return await _run(_do)


@mcp.tool()
async def drive_site_libraries(site_id: str) -> str:
    """List document libraries (drives) for a SharePoint site.

    Args:
        site_id: SharePoint site ID
    """

    def _do():
        client = get_client()
        data = client.list_site_drives(site_id)
        return format_site_drives(data)

    return await _run(_do)


@mcp.tool()
async def drive_site_list(
    site_id: str,
    drive_id: str,
    folder_id: str | None = None,
    limit: int = DEFAULT_LIMIT,
) -> str:
    """List files in a SharePoint document library.

    Args:
        site_id: SharePoint site ID
        drive_id: Document library (drive) ID
        folder_id: Folder ID within the library (omit for root)
        limit: Max items (default 20)
    """

    def _do():
        client = get_client()
        data = client.list_site_items(site_id, drive_id, folder_id=folder_id, limit=limit)
        return format_item_list(data)

    return await _run(_do)


# ======================================================================
# WRITE TOOLS
# ======================================================================


@mcp.tool()
async def drive_upload(
    file_name: str,
    content_base64: str,
    folder_path: str | None = None,
    folder_id: str | None = None,
    conflict: str = "rename",
) -> str:
    """Upload a file to OneDrive (max 4MB, simple upload).

    Requires ONEDRIVE_WRITE_ENABLED=true.

    Args:
        file_name: Name for the uploaded file
        content_base64: File content as base64-encoded string
        folder_path: Destination folder path (omit for root)
        folder_id: Destination folder ID (alternative to path)
        conflict: Conflict behavior: rename, replace, or fail (default: rename)
    """

    def _do():
        try:
            file_content = base64.b64decode(content_base64)
        except Exception:
            return "Error: Invalid base64 content"

        client = get_client()
        data = client.upload_file(
            file_content=file_content,
            file_name=file_name,
            folder_path=folder_path,
            folder_id=folder_id,
            conflict_behavior=conflict,
        )
        name = data.get("name", file_name)
        size = data.get("size", len(file_content))
        item_id = data.get("id", "")
        return f"Uploaded: {name} ({size} bytes) | id:{item_id}"

    return await _run(_do)


@mcp.tool()
async def drive_create_folder(
    name: str,
    parent_path: str | None = None,
    parent_id: str | None = None,
) -> str:
    """Create a new folder in OneDrive.

    Requires ONEDRIVE_WRITE_ENABLED=true.

    Args:
        name: Folder name
        parent_path: Parent folder path (omit for root)
        parent_id: Parent folder ID (alternative to path)
    """

    def _do():
        client = get_client()
        data = client.create_folder(name=name, parent_path=parent_path, parent_id=parent_id)
        folder_id = data.get("id", "")
        return f"Created folder: {name} | id:{folder_id}"

    return await _run(_do)


@mcp.tool()
async def drive_move(
    item_id: str,
    destination_folder_id: str | None = None,
    new_name: str | None = None,
) -> str:
    """Move and/or rename a file or folder.

    Requires ONEDRIVE_WRITE_ENABLED=true.

    Args:
        item_id: ID of the item to move/rename
        destination_folder_id: Target folder ID (for move)
        new_name: New name (for rename)
    """

    def _do():
        client = get_client()
        data = client.move_item(item_id, destination_folder_id=destination_folder_id, new_name=new_name)
        name = data.get("name", "?")
        return f"Moved/renamed: {name} | id:{data.get('id', '')}"

    return await _run(_do)


@mcp.tool()
async def drive_delete(item_id: str, confirm: bool = False) -> str:
    """Delete a file or folder (moves to recycle bin).

    Requires ONEDRIVE_WRITE_ENABLED=true AND confirm=true.

    Args:
        item_id: ID of the item to delete
        confirm: Must be true to proceed (safety gate)
    """

    def _do():
        if not confirm:
            return "Safety: pass confirm=true to delete. Item moves to recycle bin."
        client = get_client()
        client.delete_item(item_id)
        return f"Deleted item: {item_id} (moved to recycle bin)"

    return await _run(_do)


@mcp.tool()
async def drive_share(
    item_id: str,
    link_type: str = "view",
    scope: str = "organization",
    expiration: str | None = None,
) -> str:
    """Create a sharing link for a file or folder.

    Requires ONEDRIVE_WRITE_ENABLED=true.

    Args:
        item_id: ID of the item to share
        link_type: Link type: view, edit, or embed (default: view)
        scope: Sharing scope: anonymous or organization (default: organization)
        expiration: Optional ISO 8601 expiration datetime
    """

    def _do():
        client = get_client()
        data = client.create_sharing_link(item_id, link_type=link_type, scope=scope, expiration=expiration)
        return format_sharing_link(data)

    return await _run(_do)


@mcp.tool()
async def drive_permissions(item_id: str) -> str:
    """Get sharing permissions for a file or folder.

    Shows who has access and what sharing links exist.

    Args:
        item_id: ID of the item to check
    """

    def _do():
        client = get_client()
        data = client.get_permissions(item_id)
        return format_permissions(data)

    return await _run(_do)


# ======================================================================
# Entry point
# ======================================================================


def main() -> None:
    """Run the MCP server."""
    transport = os.environ.get("ONEDRIVE_MCP_TRANSPORT", "stdio").strip().lower()

    if transport == "http":
        from starlette.middleware import Middleware

        host = os.environ.get("ONEDRIVE_MCP_HOST", "127.0.0.1").strip()
        port = int(os.environ.get("ONEDRIVE_MCP_PORT", "8002").strip())

        mcp.run(
            transport="http",
            host=host,
            port=port,
            middleware=[Middleware(BearerAuthMiddleware)],
        )
    else:
        mcp.run()
