# onedrive-blade-mcp

OneDrive & SharePoint file operations via Microsoft Graph API. MCP server for the Sidereal platform.

## Features

- 18 tools: browse, read, search, upload, share, version history, delta sync, SharePoint
- Token-efficient: `$select` field filtering, pipe-delimited output, null omission
- Write-gated: destructive operations disabled by default
- Auth: device code (interactive) or client credentials (headless)
- Sidereal contract: `drive-v1`

## Quick Start

```bash
# Install
uv pip install -e .

# Configure
export ONEDRIVE_TENANT_ID="your-tenant-id"
export ONEDRIVE_CLIENT_ID="your-client-id"
export ONEDRIVE_AUTH_MODE="device_code"

# Run (stdio)
onedrive-blade-mcp

# Run (HTTP)
export ONEDRIVE_MCP_TRANSPORT=http
onedrive-blade-mcp
```

## License

MIT
