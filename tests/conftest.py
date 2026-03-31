"""Shared fixtures for onedrive-blade-mcp tests."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from onedrive_blade_mcp.client import DriveClient


@pytest.fixture()
def mock_acquire_token():
    with patch("onedrive_blade_mcp.client.acquire_token", return_value="fake-token") as m:
        yield m


@pytest.fixture()
def client(mock_acquire_token):
    c = DriveClient()
    yield c
    c.close()


@pytest.fixture()
def mock_client():
    with patch("onedrive_blade_mcp.server.get_client") as m:
        mock = MagicMock(spec=DriveClient)
        m.return_value = mock
        yield mock


# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------


@pytest.fixture()
def sample_drive_info():
    return {
        "id": "drive-001",
        "name": "OneDrive",
        "driveType": "personal",
        "owner": {"user": {"displayName": "Piers"}},
        "quota": {
            "used": 5368709120,
            "total": 1099511627776,
            "remaining": 1094142918656,
            "state": "normal",
        },
    }


@pytest.fixture()
def sample_items():
    return {
        "value": [
            {
                "id": "item-001",
                "name": "Documents",
                "size": 0,
                "lastModifiedDateTime": "2026-03-15T10:30:00Z",
                "createdDateTime": "2025-01-01T00:00:00Z",
                "folder": {"childCount": 12},
                "parentReference": {"path": "/drive/root:"},
                "webUrl": "https://onedrive.live.com/Documents",
            },
            {
                "id": "item-002",
                "name": "report.pdf",
                "size": 1048576,
                "lastModifiedDateTime": "2026-03-20T14:00:00Z",
                "createdDateTime": "2026-03-20T13:00:00Z",
                "file": {"mimeType": "application/pdf"},
                "parentReference": {"path": "/drive/root:"},
                "webUrl": "https://onedrive.live.com/report.pdf",
            },
            {
                "id": "item-003",
                "name": "photo.jpg",
                "size": 2097152,
                "lastModifiedDateTime": "2026-03-25T09:00:00Z",
                "createdDateTime": "2026-03-25T09:00:00Z",
                "file": {"mimeType": "image/jpeg"},
                "parentReference": {"path": "/drive/root:"},
                "webUrl": "https://onedrive.live.com/photo.jpg",
            },
        ]
    }


@pytest.fixture()
def sample_item_detail():
    return {
        "id": "item-002",
        "name": "report.pdf",
        "size": 1048576,
        "lastModifiedDateTime": "2026-03-20T14:00:00Z",
        "createdDateTime": "2026-03-20T13:00:00Z",
        "file": {"mimeType": "application/pdf"},
        "parentReference": {"path": "/drive/root:/Documents"},
        "webUrl": "https://onedrive.live.com/Documents/report.pdf",
        "lastModifiedBy": {"user": {"displayName": "Piers"}},
        "createdBy": {"user": {"displayName": "Piers"}},
        "description": "Quarterly financial report",
    }


@pytest.fixture()
def sample_versions():
    return {
        "value": [
            {
                "id": "1.0",
                "lastModifiedDateTime": "2026-03-15T10:00:00Z",
                "lastModifiedBy": {"user": {"displayName": "Piers"}},
                "size": 512000,
            },
            {
                "id": "2.0",
                "lastModifiedDateTime": "2026-03-20T14:00:00Z",
                "lastModifiedBy": {"user": {"displayName": "Piers"}},
                "size": 1048576,
            },
        ]
    }


@pytest.fixture()
def sample_delta():
    return {
        "value": [
            {
                "id": "item-002",
                "name": "report.pdf",
                "lastModifiedDateTime": "2026-03-20T14:00:00Z",
                "file": {"mimeType": "application/pdf"},
            },
            {
                "id": "item-004",
                "name": "deleted.txt",
                "lastModifiedDateTime": "2026-03-19T12:00:00Z",
                "deleted": {"state": "deleted"},
            },
        ],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/drive/root/delta?token=abc123",
    }


@pytest.fixture()
def sample_sites():
    return {
        "value": [
            {
                "id": "site-001",
                "name": "marketing",
                "displayName": "Marketing",
                "webUrl": "https://contoso.sharepoint.com/sites/marketing",
                "description": "Marketing team site",
            },
            {
                "id": "site-002",
                "name": "engineering",
                "displayName": "Engineering",
                "webUrl": "https://contoso.sharepoint.com/sites/engineering",
                "description": "",
            },
        ]
    }


@pytest.fixture()
def sample_permissions():
    return {
        "value": [
            {
                "id": "perm-001",
                "roles": ["read"],
                "grantedToV2": {"user": {"displayName": "Alice", "email": "alice@contoso.com"}},
            },
            {
                "id": "perm-002",
                "roles": ["write"],
                "link": {
                    "type": "edit",
                    "webUrl": "https://onedrive.live.com/share/abc123",
                },
            },
        ]
    }
