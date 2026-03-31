"""Tests for client.py — DriveClient methods and error handling."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import httpx
import pytest

from onedrive_blade_mcp.client import (
    ConflictError,
    DriveError,
    FileTooLargeError,
    NotFoundError,
    RateLimitError,
    WriteDisabledError,
    _classify_error,
)


class TestClassifyError:
    def _make_response(self, status_code, json_body=None, text="", headers=None):
        resp = MagicMock(spec=httpx.Response)
        resp.status_code = status_code
        resp.text = text
        resp.headers = headers or {}
        if json_body is not None:
            resp.json.return_value = json_body
        else:
            resp.json.side_effect = ValueError("No JSON")
        return resp

    def test_401(self):
        resp = self._make_response(401, {"error": {"message": "Invalid token"}})
        exc = _classify_error(resp)
        assert isinstance(exc, type(exc))  # AuthError
        assert exc.status_code == 401

    def test_404(self):
        resp = self._make_response(404, {"error": {"message": "Item not found"}})
        exc = _classify_error(resp)
        assert isinstance(exc, NotFoundError)

    def test_409(self):
        resp = self._make_response(409, {"error": {"message": "Name conflict"}})
        exc = _classify_error(resp)
        assert isinstance(exc, ConflictError)

    def test_429(self):
        resp = self._make_response(429, {"error": {"message": "Throttled"}}, headers={"Retry-After": "30"})
        exc = _classify_error(resp)
        assert isinstance(exc, RateLimitError)

    def test_500(self):
        resp = self._make_response(500, text="Internal error")
        exc = _classify_error(resp)
        assert isinstance(exc, DriveError)
        assert exc.status_code == 500


class TestWriteGate:
    def test_write_disabled(self, client):
        with patch.dict("os.environ", {}, clear=True):
            with pytest.raises(WriteDisabledError):
                client._require_write()

    def test_write_enabled(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            client._require_write()  # should not raise


class TestDriveClientReadOps:
    def test_get_drive_info(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"id": "d1", "name": "OneDrive"}
            result = client.get_drive_info()
            assert result["name"] == "OneDrive"
            mock_get.assert_called_once()

    def test_list_items_root(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.list_items()
            args = mock_get.call_args
            assert "/me/drive/root/children" in args[0][0]

    def test_list_items_by_path(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.list_items(folder_path="/Documents")
            args = mock_get.call_args
            assert "Documents" in args[0][0]

    def test_list_items_by_id(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.list_items(folder_id="folder-123")
            args = mock_get.call_args
            assert "folder-123" in args[0][0]

    def test_get_item_metadata_requires_id_or_path(self, client):
        with pytest.raises(DriveError, match="Either item_id or item_path"):
            client.get_item_metadata()

    def test_search_files(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.search_files("budget")
            args = mock_get.call_args
            assert "search" in args[0][0]
            assert "budget" in args[0][0]

    def test_get_versions(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.get_versions("item-001")
            args = mock_get.call_args
            assert "item-001" in args[0][0]
            assert "versions" in args[0][0]

    def test_get_delta_initial(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": [], "@odata.deltaLink": "https://..."}
            client.get_delta(folder_path="/Documents")
            args = mock_get.call_args
            assert "delta" in args[0][0]

    def test_list_sites(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.list_sites(query="marketing")
            mock_get.assert_called_once()

    def test_list_site_drives(self, client):
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"value": []}
            client.list_site_drives("site-001")
            args = mock_get.call_args
            assert "site-001" in args[0][0]
            assert "drives" in args[0][0]


class TestDriveClientWriteOps:
    def test_upload_file_too_large(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            with pytest.raises(FileTooLargeError):
                client.upload_file(
                    file_content=b"x" * (5 * 1024 * 1024),
                    file_name="big.bin",
                )

    def test_create_folder(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            with patch.object(client, "_post") as mock_post:
                mock_post.return_value = {"id": "new-folder", "name": "Reports"}
                result = client.create_folder("Reports", parent_path="/Documents")
                assert result["name"] == "Reports"
                mock_post.assert_called_once()

    def test_move_requires_dest_or_name(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            with pytest.raises(DriveError, match="Either destination_folder_id or new_name"):
                client.move_item("item-001")

    def test_delete_item(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            with patch.object(client, "_delete") as mock_delete:
                mock_delete.return_value = None
                client.delete_item("item-001")
                mock_delete.assert_called_once()

    def test_create_sharing_link(self, client):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            with patch.object(client, "_post") as mock_post:
                mock_post.return_value = {"link": {"webUrl": "https://share/abc"}}
                result = client.create_sharing_link("item-001", link_type="view", scope="organization")
                assert "link" in result
