"""Tests for models.py — write gate, scopes, constants."""

from __future__ import annotations

from unittest.mock import patch

from onedrive_blade_mcp.models import (
    DEFAULT_LIMIT,
    GRAPH_BASE_URL,
    MAX_UPLOAD_SIZE,
    SCOPES_CLIENT_CREDENTIALS,
    SCOPES_READ,
    SCOPES_READWRITE,
    get_scopes,
    is_write_enabled,
    require_write,
)


class TestWriteGate:
    def test_disabled_by_default(self):
        with patch.dict("os.environ", {}, clear=True):
            assert is_write_enabled() is False

    def test_enabled(self):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            assert is_write_enabled() is True

    def test_case_insensitive(self):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "True"}):
            assert is_write_enabled() is True  # .lower() comparison

    def test_other_value(self):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "yes"}):
            assert is_write_enabled() is False

    def test_require_write_disabled(self):
        with patch.dict("os.environ", {}, clear=True):
            msg = require_write()
            assert msg is not None
            assert "disabled" in msg.lower()

    def test_require_write_enabled(self):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            assert require_write() is None


class TestScopes:
    def test_read_scopes(self):
        with patch.dict("os.environ", {}, clear=True):
            scopes = get_scopes()
            assert scopes == SCOPES_READ
            assert "Files.Read" in scopes

    def test_write_scopes(self):
        with patch.dict("os.environ", {"ONEDRIVE_WRITE_ENABLED": "true"}):
            scopes = get_scopes()
            assert scopes == SCOPES_READWRITE
            assert "Files.ReadWrite" in scopes


class TestConstants:
    def test_default_limit(self):
        assert DEFAULT_LIMIT == 20

    def test_max_upload_size(self):
        assert MAX_UPLOAD_SIZE == 4 * 1024 * 1024

    def test_graph_base_url(self):
        assert GRAPH_BASE_URL == "https://graph.microsoft.com/v1.0"

    def test_client_credentials_scopes(self):
        assert SCOPES_CLIENT_CREDENTIALS == ["https://graph.microsoft.com/.default"]
