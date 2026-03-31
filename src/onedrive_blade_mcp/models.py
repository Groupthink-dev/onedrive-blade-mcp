"""Shared constants, types, and write-gate for OneDrive Blade MCP server."""

from __future__ import annotations

import logging
import os

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_LIMIT = 20
MAX_UPLOAD_SIZE = 4 * 1024 * 1024  # 4 MB (simple upload limit; larger needs resumable)
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Graph API field selections (token efficiency)
# ---------------------------------------------------------------------------

ITEM_LIST_FIELDS = [
    "id",
    "name",
    "size",
    "lastModifiedDateTime",
    "createdDateTime",
    "file",
    "folder",
    "parentReference",
    "webUrl",
]

ITEM_DETAIL_FIELDS = ITEM_LIST_FIELDS + [
    "lastModifiedBy",
    "createdBy",
    "description",
    "shared",
    "permissions",
]

# ---------------------------------------------------------------------------
# Permission scopes
# ---------------------------------------------------------------------------

SCOPES_READ = [
    "User.Read",
    "Files.Read",
    "Sites.Read.All",
]

SCOPES_READWRITE = [
    "User.Read",
    "Files.ReadWrite",
    "Sites.ReadWrite.All",
]

SCOPES_CLIENT_CREDENTIALS = [
    "https://graph.microsoft.com/.default",
]


def get_scopes() -> list[str]:
    """Return scopes based on write mode."""
    if is_write_enabled():
        return SCOPES_READWRITE
    return SCOPES_READ


# ---------------------------------------------------------------------------
# Write gate
# ---------------------------------------------------------------------------


def is_write_enabled() -> bool:
    """Check if write operations are enabled via env var."""
    return os.environ.get("ONEDRIVE_WRITE_ENABLED", "").lower() == "true"


def require_write() -> str | None:
    """Return an error message if writes are disabled, else None."""
    if not is_write_enabled():
        return "Error: Write operations are disabled. Set ONEDRIVE_WRITE_ENABLED=true to enable."
    return None
