"""Token-efficient formatters for OneDrive & SharePoint data.

Same conventions as other blade MCPs: pipe-delimited, null omission, capped lists.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _compact_datetime(iso: str | None) -> str:
    """Shorten ISO datetime to 'Mar 15 14:30' format."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%b %d %H:%M")
    except (ValueError, TypeError):
        return iso[:16] if iso else ""


def _human_size(size: int | None) -> str:
    """Format bytes as human-readable string."""
    if size is None:
        return ""
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(size) < 1024:
            if unit == "B":
                return f"{size}{unit}"
            return f"{size:.1f}{unit}"
        size /= 1024  # type: ignore[assignment]
    return f"{size:.1f}PB"


def _item_type_icon(item: dict[str, Any]) -> str:
    """Return a compact type indicator for a drive item."""
    if "folder" in item:
        child_count = item["folder"].get("childCount", 0)
        return f"[dir:{child_count}]"
    if "file" in item:
        mime = item.get("file", {}).get("mimeType", "")
        if mime.startswith("image/"):
            return "[img]"
        if mime.startswith("video/"):
            return "[vid]"
        if mime.startswith("audio/"):
            return "[aud]"
        if "pdf" in mime:
            return "[pdf]"
        if "spreadsheet" in mime or "excel" in mime:
            return "[xls]"
        if "document" in mime or "word" in mime:
            return "[doc]"
        if "presentation" in mime or "powerpoint" in mime:
            return "[ppt]"
        return "[file]"
    return "[?]"


def _truncate(s: str, max_len: int = 120) -> str:
    if len(s) <= max_len:
        return s
    return s[: max_len - 3] + "..."


# ---------------------------------------------------------------------------
# Drive info
# ---------------------------------------------------------------------------


def format_drive_info(drive: dict[str, Any]) -> str:
    """Format drive metadata."""
    lines = []
    lines.append(f"Drive: {drive.get('name', 'OneDrive')} ({drive.get('driveType', 'personal')})")

    quota = drive.get("quota", {})
    if quota:
        used = _human_size(quota.get("used"))
        total = _human_size(quota.get("total"))
        remaining = _human_size(quota.get("remaining"))
        if used and total:
            lines.append(f"Usage: {used} / {total} ({remaining} free)")

        state = quota.get("state", "")
        if state and state != "normal":
            lines.append(f"State: {state}")

    owner = drive.get("owner", {})
    user = owner.get("user", {})
    if user.get("displayName"):
        lines.append(f"Owner: {user['displayName']}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Item lists
# ---------------------------------------------------------------------------


def format_item_list(data: dict[str, Any], folder_label: str | None = None) -> str:
    """Format a list of drive items (folder children or search results)."""
    items = data.get("value", [])
    if not items:
        return f"Empty folder: {folder_label}" if folder_label else "No items found."

    lines = []
    if folder_label:
        lines.append(f"📁 {folder_label} ({len(items)} items)")
        lines.append("")

    for item in items:
        icon = _item_type_icon(item)
        name = item.get("name", "?")
        size = _human_size(item.get("size"))
        modified = _compact_datetime(item.get("lastModifiedDateTime"))
        item_id = item.get("id", "")

        parts = [icon, name]
        if size and "folder" not in item:
            parts.append(size)
        if modified:
            parts.append(modified)
        parts.append(f"id:{item_id}")

        lines.append(" | ".join(parts))

    # Pagination
    next_link = data.get("@odata.nextLink")
    total = data.get("@odata.count")
    if next_link:
        shown = len(items)
        if total:
            lines.append(f"\n--- Showing {shown} of {total} items (more available) ---")
        else:
            lines.append(f"\n--- Showing {shown} items (more available) ---")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Item detail
# ---------------------------------------------------------------------------


def format_item_detail(item: dict[str, Any]) -> str:
    """Format detailed metadata for a single item."""
    lines = []
    icon = _item_type_icon(item)
    name = item.get("name", "?")
    lines.append(f"{icon} {name}")

    item_id = item.get("id", "")
    if item_id:
        lines.append(f"ID: {item_id}")

    size = item.get("size")
    if size is not None:
        lines.append(f"Size: {_human_size(size)}")

    created = item.get("createdDateTime")
    if created:
        lines.append(f"Created: {_compact_datetime(created)}")

    modified = item.get("lastModifiedDateTime")
    if modified:
        lines.append(f"Modified: {_compact_datetime(modified)}")

    # Created/modified by
    created_by = item.get("createdBy", {}).get("user", {}).get("displayName")
    if created_by:
        lines.append(f"Created by: {created_by}")

    modified_by = item.get("lastModifiedBy", {}).get("user", {}).get("displayName")
    if modified_by:
        lines.append(f"Modified by: {modified_by}")

    # Parent path
    parent = item.get("parentReference", {})
    parent_path = parent.get("path")
    if parent_path:
        # Strip the /drive/root: prefix for readability
        display_path = parent_path.replace("/drive/root:", "")
        lines.append(f"Path: {display_path}/{name}")

    web_url = item.get("webUrl")
    if web_url:
        lines.append(f"URL: {web_url}")

    description = item.get("description")
    if description:
        lines.append(f"Description: {_truncate(description)}")

    # Folder-specific
    if "folder" in item:
        child_count = item["folder"].get("childCount", 0)
        lines.append(f"Children: {child_count}")

    # File-specific
    if "file" in item:
        mime = item["file"].get("mimeType", "")
        if mime:
            lines.append(f"MIME: {mime}")

    # Sharing
    shared = item.get("shared")
    if shared:
        scope = shared.get("scope", "unknown")
        shared_by = shared.get("sharedBy", {}).get("user", {}).get("displayName", "")
        lines.append(f"Shared: {scope}" + (f" by {shared_by}" if shared_by else ""))

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Search results
# ---------------------------------------------------------------------------


def format_search_results(data: dict[str, Any], query: str) -> str:
    """Format search results with query context."""
    items = data.get("value", [])
    if not items:
        return f"No results for: {query}"

    lines = [f"Search results for '{query}' ({len(items)} matches)", ""]

    for item in items:
        icon = _item_type_icon(item)
        name = item.get("name", "?")
        size = _human_size(item.get("size"))
        modified = _compact_datetime(item.get("lastModifiedDateTime"))
        item_id = item.get("id", "")

        # Include parent path for search context
        parent = item.get("parentReference", {})
        parent_path = parent.get("path", "")
        display_path = parent_path.replace("/drive/root:", "") if parent_path else ""

        parts = [icon, name]
        if display_path:
            parts.append(f"in:{display_path}")
        if size and "folder" not in item:
            parts.append(size)
        if modified:
            parts.append(modified)
        parts.append(f"id:{item_id}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Versions
# ---------------------------------------------------------------------------


def format_versions(data: dict[str, Any], item_name: str | None = None) -> str:
    """Format version history."""
    versions = data.get("value", [])
    if not versions:
        return "No version history available."

    header = f"Versions for {item_name}" if item_name else "Version history"
    lines = [f"{header} ({len(versions)} versions)", ""]

    for v in versions:
        vid = v.get("id", "?")
        modified = _compact_datetime(v.get("lastModifiedDateTime"))
        size = _human_size(v.get("size"))
        modified_by = v.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")

        parts = [f"v{vid}"]
        if modified:
            parts.append(modified)
        if modified_by:
            parts.append(modified_by)
        if size:
            parts.append(size)

        lines.append(" | ".join(parts))

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Delta (change tracking)
# ---------------------------------------------------------------------------


def format_delta(data: dict[str, Any]) -> str:
    """Format delta query results."""
    items = data.get("value", [])
    delta_link = data.get("@odata.deltaLink")
    next_link = data.get("@odata.nextLink")

    if not items and delta_link:
        return f"No changes detected.\n\ndelta_link: {delta_link}"

    lines = [f"Changes detected: {len(items)}", ""]

    for item in items:
        deleted = item.get("deleted")
        icon = "[DEL]" if deleted else _item_type_icon(item)
        name = item.get("name", "?")
        modified = _compact_datetime(item.get("lastModifiedDateTime"))
        item_id = item.get("id", "")

        parts = [icon, name]
        if modified:
            parts.append(modified)
        parts.append(f"id:{item_id}")

        lines.append(" | ".join(parts))

    lines.append("")
    if delta_link:
        lines.append(f"delta_link: {delta_link}")
    elif next_link:
        lines.append(f"next_link: {next_link}")
        lines.append("(More changes available — pass next_link as delta_link)")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Sites
# ---------------------------------------------------------------------------


def format_sites(data: dict[str, Any]) -> str:
    """Format SharePoint sites list."""
    sites = data.get("value", [])
    if not sites:
        return "No SharePoint sites found."

    lines = [f"SharePoint sites ({len(sites)})", ""]

    for site in sites:
        name = site.get("displayName") or site.get("name", "?")
        site_id = site.get("id", "")
        url = site.get("webUrl", "")
        desc = site.get("description", "")

        parts = [name]
        if desc:
            parts.append(_truncate(desc, 60))
        if url:
            parts.append(url)
        parts.append(f"id:{site_id}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


def format_site_drives(data: dict[str, Any], site_name: str | None = None) -> str:
    """Format document libraries for a site."""
    drives = data.get("value", [])
    if not drives:
        return "No document libraries found."

    header = f"Document libraries for {site_name}" if site_name else "Document libraries"
    lines = [f"{header} ({len(drives)})", ""]

    for drive in drives:
        name = drive.get("name", "?")
        drive_id = drive.get("id", "")
        drive_type = drive.get("driveType", "")
        url = drive.get("webUrl", "")

        quota = drive.get("quota", {})
        used = _human_size(quota.get("used")) if quota.get("used") else ""

        parts = [name]
        if drive_type:
            parts.append(drive_type)
        if used:
            parts.append(used)
        if url:
            parts.append(url)
        parts.append(f"id:{drive_id}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Permissions & sharing
# ---------------------------------------------------------------------------


def format_permissions(data: dict[str, Any]) -> str:
    """Format permissions list for an item."""
    perms = data.get("value", [])
    if not perms:
        return "No sharing permissions set."

    lines = [f"Permissions ({len(perms)})", ""]

    for p in perms:
        perm_id = p.get("id", "?")
        roles = ", ".join(p.get("roles", []))

        # Identify the grantee
        granted_to = p.get("grantedToV2") or p.get("grantedTo", {})
        user = granted_to.get("user", {})
        grantee = user.get("displayName") or user.get("email", "")

        # Link info
        link = p.get("link", {})
        link_type = link.get("type", "")
        link_url = link.get("webUrl", "")

        parts = []
        if grantee:
            parts.append(grantee)
        elif link_type:
            parts.append(f"link:{link_type}")
        parts.append(f"roles:{roles}")
        if link_url:
            parts.append(link_url)
        parts.append(f"id:{perm_id}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


def format_sharing_link(data: dict[str, Any]) -> str:
    """Format a newly created sharing link."""
    link = data.get("link", {})
    url = link.get("webUrl", "")
    link_type = link.get("type", "")
    scope = link.get("scope", "")
    expiration = data.get("expirationDateTime")

    lines = ["Sharing link created:"]
    lines.append(f"URL: {url}")
    lines.append(f"Type: {link_type} | Scope: {scope}")
    if expiration:
        lines.append(f"Expires: {_compact_datetime(expiration)}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# File content
# ---------------------------------------------------------------------------


def format_file_content(data: dict[str, Any]) -> str:
    """Format file content read result."""
    meta = data.get("metadata", {})
    name = meta.get("name", "?")
    size = _human_size(meta.get("size"))
    content_type = data.get("content_type", "unknown")
    truncated = data.get("truncated", False)

    lines = [f"File: {name} ({size})"]

    if truncated:
        lines.append("⚠ Content truncated (file too large to read fully)")

    if content_type == "text":
        content = data.get("content", "")
        lines.append("")
        lines.append(content)
    elif content_type == "binary":
        mime = data.get("mime_type", "binary")
        b64 = data.get("content_base64", "")
        b64_size = len(b64) if b64 else 0
        lines.append(f"Binary content ({mime}, {b64_size} bytes base64)")
        if b64:
            lines.append(f"\nbase64: {_truncate(b64, 200)}")
    else:
        lines.append("Unknown content type")

    return "\n".join(lines)
