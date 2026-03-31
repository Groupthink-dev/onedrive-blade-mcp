"""Tests for formatters.py — token-efficient output formatting."""

from __future__ import annotations

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


class TestFormatDriveInfo:
    def test_basic(self, sample_drive_info):
        result = format_drive_info(sample_drive_info)
        assert "OneDrive" in result
        assert "personal" in result
        assert "Piers" in result
        assert "free" in result

    def test_no_quota(self):
        result = format_drive_info({"name": "Test", "driveType": "business"})
        assert "Test" in result
        assert "business" in result


class TestFormatItemList:
    def test_with_items(self, sample_items):
        result = format_item_list(sample_items, folder_label="Root")
        assert "Root" in result
        assert "3 items" in result
        assert "Documents" in result
        assert "report.pdf" in result
        assert "[dir:12]" in result
        assert "[pdf]" in result
        assert "[img]" in result

    def test_empty(self):
        result = format_item_list({"value": []}, folder_label="Empty")
        assert "Empty" in result

    def test_no_label(self, sample_items):
        result = format_item_list(sample_items)
        assert "Documents" in result

    def test_pagination(self, sample_items):
        sample_items["@odata.nextLink"] = "https://graph.microsoft.com/..."
        result = format_item_list(sample_items)
        assert "more available" in result


class TestFormatItemDetail:
    def test_full(self, sample_item_detail):
        result = format_item_detail(sample_item_detail)
        assert "report.pdf" in result
        assert "[pdf]" in result
        assert "item-002" in result
        assert "Piers" in result
        assert "Quarterly financial" in result
        assert "application/pdf" in result

    def test_folder(self):
        item = {
            "id": "folder-1",
            "name": "Projects",
            "folder": {"childCount": 5},
            "lastModifiedDateTime": "2026-03-15T10:00:00Z",
        }
        result = format_item_detail(item)
        assert "Projects" in result
        assert "[dir:5]" in result
        assert "Children: 5" in result


class TestFormatSearchResults:
    def test_with_results(self, sample_items):
        result = format_search_results(sample_items, "report")
        assert "report" in result
        assert "3 matches" in result

    def test_no_results(self):
        result = format_search_results({"value": []}, "nonexistent")
        assert "No results" in result
        assert "nonexistent" in result


class TestFormatVersions:
    def test_with_versions(self, sample_versions):
        result = format_versions(sample_versions, item_name="report.pdf")
        assert "report.pdf" in result
        assert "2 versions" in result
        assert "v1.0" in result
        assert "v2.0" in result
        assert "Piers" in result

    def test_empty(self):
        result = format_versions({"value": []})
        assert "No version history" in result


class TestFormatDelta:
    def test_with_changes(self, sample_delta):
        result = format_delta(sample_delta)
        assert "2" in result
        assert "report.pdf" in result
        assert "[DEL]" in result
        assert "deleted.txt" in result
        assert "delta_link" in result

    def test_no_changes(self):
        data = {"value": [], "@odata.deltaLink": "https://example.com/delta?token=abc"}
        result = format_delta(data)
        assert "No changes" in result
        assert "delta_link" in result


class TestFormatSites:
    def test_with_sites(self, sample_sites):
        result = format_sites(sample_sites)
        assert "2" in result
        assert "Marketing" in result
        assert "Engineering" in result

    def test_empty(self):
        result = format_sites({"value": []})
        assert "No SharePoint sites" in result


class TestFormatSiteDrives:
    def test_with_drives(self):
        data = {
            "value": [
                {
                    "id": "drive-1",
                    "name": "Documents",
                    "driveType": "documentLibrary",
                    "webUrl": "https://contoso.sharepoint.com/Documents",
                    "quota": {"used": 1073741824},
                },
            ]
        }
        result = format_site_drives(data, site_name="Marketing")
        assert "Marketing" in result
        assert "Documents" in result
        assert "documentLibrary" in result

    def test_empty(self):
        result = format_site_drives({"value": []})
        assert "No document libraries" in result


class TestFormatPermissions:
    def test_with_permissions(self, sample_permissions):
        result = format_permissions(sample_permissions)
        assert "2" in result
        assert "Alice" in result
        assert "read" in result
        assert "link:edit" in result

    def test_empty(self):
        result = format_permissions({"value": []})
        assert "No sharing permissions" in result


class TestFormatSharingLink:
    def test_basic(self):
        data = {
            "link": {
                "webUrl": "https://onedrive.live.com/share/xyz",
                "type": "view",
                "scope": "organization",
            },
        }
        result = format_sharing_link(data)
        assert "https://onedrive.live.com/share/xyz" in result
        assert "view" in result
        assert "organization" in result

    def test_with_expiration(self):
        data = {
            "link": {"webUrl": "https://example.com", "type": "edit", "scope": "anonymous"},
            "expirationDateTime": "2026-04-15T23:59:59Z",
        }
        result = format_sharing_link(data)
        assert "Expires" in result


class TestFormatFileContent:
    def test_text_file(self):
        data = {
            "metadata": {"name": "notes.txt", "size": 42},
            "content_type": "text",
            "content": "Hello, world!",
            "truncated": False,
        }
        result = format_file_content(data)
        assert "notes.txt" in result
        assert "Hello, world!" in result

    def test_truncated(self):
        data = {
            "metadata": {"name": "big.log", "size": 50_000_000},
            "content_type": "text",
            "content": "start...",
            "truncated": True,
        }
        result = format_file_content(data)
        assert "truncated" in result.lower()

    def test_binary(self):
        data = {
            "metadata": {"name": "image.png", "size": 1024},
            "content_type": "binary",
            "content_base64": "iVBORw0KGgo=",
            "truncated": False,
            "mime_type": "image/png",
        }
        result = format_file_content(data)
        assert "Binary" in result
        assert "image/png" in result
