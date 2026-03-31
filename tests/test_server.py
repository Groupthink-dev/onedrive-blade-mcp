"""Tests for server.py — MCP tool handlers."""

from __future__ import annotations

import base64

import pytest

from onedrive_blade_mcp.client import DriveError, NotFoundError, WriteDisabledError
from onedrive_blade_mcp.server import (
    drive_create_folder,
    drive_delete,
    drive_delta,
    drive_download,
    drive_info,
    drive_list,
    drive_metadata,
    drive_move,
    drive_permissions,
    drive_read,
    drive_search,
    drive_share,
    drive_site_libraries,
    drive_site_list,
    drive_sites,
    drive_thumbnail,
    drive_upload,
    drive_versions,
)

# ---------------------------------------------------------------------------
# Meta
# ---------------------------------------------------------------------------


class TestDriveInfo:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_drive_info):
        mock_client.get_drive_info.return_value = sample_drive_info
        result = await drive_info()
        assert "OneDrive" in result
        assert "personal" in result
        mock_client.get_drive_info.assert_called_once()


# ---------------------------------------------------------------------------
# Read tools
# ---------------------------------------------------------------------------


class TestDriveList:
    @pytest.mark.asyncio
    async def test_root(self, mock_client, sample_items):
        mock_client.list_items.return_value = sample_items
        result = await drive_list()
        assert "Root" in result
        assert "Documents" in result
        mock_client.list_items.assert_called_once_with(folder_path=None, folder_id=None, limit=20)

    @pytest.mark.asyncio
    async def test_with_path(self, mock_client, sample_items):
        mock_client.list_items.return_value = sample_items
        result = await drive_list(folder_path="/Documents")
        assert "Documents" in result

    @pytest.mark.asyncio
    async def test_error(self, mock_client):
        mock_client.list_items.side_effect = DriveError("Not found", 404)
        result = await drive_list(folder_path="/nonexistent")
        assert "Error" in result


class TestDriveRead:
    @pytest.mark.asyncio
    async def test_text_file(self, mock_client):
        mock_client.read_file_content.return_value = {
            "metadata": {"name": "test.txt", "size": 13},
            "content_type": "text",
            "content": "Hello, world!",
            "truncated": False,
        }
        result = await drive_read(item_path="/test.txt")
        assert "Hello, world!" in result

    @pytest.mark.asyncio
    async def test_not_found(self, mock_client):
        mock_client.read_file_content.side_effect = NotFoundError("Not found", 404)
        result = await drive_read(item_path="/missing.txt")
        assert "Error" in result


class TestDriveMetadata:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_item_detail):
        mock_client.get_item_metadata.return_value = sample_item_detail
        result = await drive_metadata(item_path="/Documents/report.pdf")
        assert "report.pdf" in result
        assert "Piers" in result


class TestDriveSearch:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_items):
        mock_client.search_files.return_value = sample_items
        result = await drive_search(query="report")
        assert "report" in result
        assert "3 matches" in result

    @pytest.mark.asyncio
    async def test_no_results(self, mock_client):
        mock_client.search_files.return_value = {"value": []}
        result = await drive_search(query="nonexistent")
        assert "No results" in result


class TestDriveDownload:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.download_url.return_value = "https://download.example.com/file"
        result = await drive_download(item_id="item-001")
        assert "https://download.example.com/file" in result


class TestDriveVersions:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_versions):
        mock_client.get_versions.return_value = sample_versions
        result = await drive_versions(item_id="item-002")
        assert "2 versions" in result
        assert "v1.0" in result


class TestDriveDelta:
    @pytest.mark.asyncio
    async def test_initial(self, mock_client, sample_delta):
        mock_client.get_delta.return_value = sample_delta
        result = await drive_delta(folder_path="/Documents")
        assert "2" in result  # changes detected
        assert "delta_link" in result

    @pytest.mark.asyncio
    async def test_with_link(self, mock_client):
        mock_client.get_delta.return_value = {
            "value": [],
            "@odata.deltaLink": "https://graph.microsoft.com/delta?token=new",
        }
        result = await drive_delta(delta_link="https://graph.microsoft.com/delta?token=old")
        assert "No changes" in result


class TestDriveThumbnail:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.get_thumbnail.return_value = {
            "value": [
                {
                    "small": {"width": 96, "height": 96, "url": "https://thumb/s"},
                    "medium": {"width": 176, "height": 176, "url": "https://thumb/m"},
                    "large": {"width": 800, "height": 800, "url": "https://thumb/l"},
                }
            ]
        }
        result = await drive_thumbnail(item_id="item-003")
        assert "96x96" in result
        assert "https://thumb/m" in result

    @pytest.mark.asyncio
    async def test_no_thumbnails(self, mock_client):
        mock_client.get_thumbnail.return_value = {"value": []}
        result = await drive_thumbnail(item_id="item-001")
        assert "No thumbnails" in result


class TestDriveSites:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_sites):
        mock_client.list_sites.return_value = sample_sites
        result = await drive_sites()
        assert "Marketing" in result
        assert "Engineering" in result


class TestDriveSiteLibraries:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.list_site_drives.return_value = {
            "value": [
                {"id": "d1", "name": "Documents", "driveType": "documentLibrary", "webUrl": "https://sp/docs"}
            ]
        }
        result = await drive_site_libraries(site_id="site-001")
        assert "Documents" in result


class TestDriveSiteList:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_items):
        mock_client.list_site_items.return_value = sample_items
        result = await drive_site_list(site_id="site-001", drive_id="drive-001")
        assert "Documents" in result


# ---------------------------------------------------------------------------
# Write tools
# ---------------------------------------------------------------------------


class TestDriveUpload:
    @pytest.mark.asyncio
    async def test_write_disabled(self, mock_client):
        mock_client.upload_file.side_effect = WriteDisabledError("Writes disabled")
        content = base64.b64encode(b"hello").decode()
        result = await drive_upload(file_name="test.txt", content_base64=content)
        assert "disabled" in result.lower()

    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.upload_file.return_value = {"id": "new-001", "name": "test.txt", "size": 5}
        content = base64.b64encode(b"hello").decode()
        result = await drive_upload(file_name="test.txt", content_base64=content, folder_path="/Documents")
        assert "Uploaded" in result
        assert "test.txt" in result

    @pytest.mark.asyncio
    async def test_invalid_base64(self, mock_client):
        result = await drive_upload(file_name="test.txt", content_base64="not-valid-base64!!!")
        assert "Error" in result


class TestDriveCreateFolder:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.create_folder.return_value = {"id": "folder-new", "name": "Reports"}
        result = await drive_create_folder(name="Reports", parent_path="/Documents")
        assert "Created folder" in result
        assert "Reports" in result


class TestDriveMove:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.move_item.return_value = {"id": "item-001", "name": "moved.txt"}
        result = await drive_move(item_id="item-001", destination_folder_id="folder-002")
        assert "Moved" in result


class TestDriveDelete:
    @pytest.mark.asyncio
    async def test_requires_confirm(self, mock_client):
        result = await drive_delete(item_id="item-001")
        assert "confirm=true" in result
        mock_client.delete_item.assert_not_called()

    @pytest.mark.asyncio
    async def test_success_with_confirm(self, mock_client):
        mock_client.delete_item.return_value = None
        result = await drive_delete(item_id="item-001", confirm=True)
        assert "Deleted" in result
        mock_client.delete_item.assert_called_once_with("item-001")


class TestDriveShare:
    @pytest.mark.asyncio
    async def test_success(self, mock_client):
        mock_client.create_sharing_link.return_value = {
            "link": {"webUrl": "https://share.example.com/abc", "type": "view", "scope": "organization"},
        }
        result = await drive_share(item_id="item-001")
        assert "https://share.example.com/abc" in result
        assert "view" in result


class TestDrivePermissions:
    @pytest.mark.asyncio
    async def test_success(self, mock_client, sample_permissions):
        mock_client.get_permissions.return_value = sample_permissions
        result = await drive_permissions(item_id="item-001")
        assert "Alice" in result
        assert "read" in result
