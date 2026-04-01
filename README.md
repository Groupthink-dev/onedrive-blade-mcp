# onedrive-blade-mcp

A token-efficient [MCP](https://modelcontextprotocol.io) server for OneDrive and SharePoint file operations via Microsoft Graph API.

18 tools covering browse, read, search, upload, share, versioning, delta sync, and SharePoint — with pipe-delimited output, write-gating, and dual auth.

## Why this over generic file MCPs?

| | **onedrive-blade-mcp** | Generic file / Graph MCPs |
|---|---|---|
| **Output** | Pipe-delimited, null omission | Raw JSON blobs (10-50x more tokens) |
| **Delta sync** | `drive_delta` tracks file changes incrementally | Full re-listing every time |
| **SharePoint** | Browse sites, libraries, and files natively | OneDrive-only or not supported |
| **Safety** | Write-gated by default, delete requires `confirm=true` | Varies |
| **Auth** | Dual: device code (interactive) + client credentials (headless/CI) | Single flow |
| **Versioning** | `drive_versions` for full file history | Not exposed |
| **Thumbnails** | `drive_thumbnail` for image previews | Not available |

Delta sync is the killer feature for AI workflows — track what changed in a folder without re-listing thousands of files.

## Features

- **18 tools** — browse, read, search, upload, share, version history, delta sync, SharePoint
- **Token-efficient output** — pipe-delimited, field-selected, null fields omitted
- **Write-gated by default** — reads are free, writes require `ONEDRIVE_WRITE_ENABLED=true`
- **Delta sync** — incremental file change tracking with delta tokens
- **SharePoint integration** — sites, document libraries, and files
- **Dual authentication** — device code for interactive, client credentials for headless
- **Sharing management** — create links, check permissions, control scope
- **FastMCP 2.0** — stdio and HTTP transports

## Quick Start

```bash
# Install
git clone https://github.com/groupthink-dev/onedrive-blade-mcp.git
cd onedrive-blade-mcp
uv sync

# Configure
export ONEDRIVE_TENANT_ID="your-tenant-id"
export ONEDRIVE_CLIENT_ID="your-client-id"
export ONEDRIVE_AUTH_MODE="device_code"

# Run (stdio)
uv run onedrive-blade-mcp
```

## Tools (18)

### Meta

| Tool | Description |
|------|-------------|
| `drive_info` | Storage info — drive type, quota usage, owner |

### Read (8)

| Tool | Description |
|------|-------------|
| `drive_list` | Browse folder contents (path or ID) |
| `drive_read` | Read file content — text for text files, base64 summary for binary |
| `drive_metadata` | Detailed metadata — name, size, dates, creator, MIME, sharing |
| `drive_search` | Search files by name or content (Microsoft Search) |
| `drive_download` | Pre-authenticated short-lived download URL |
| `drive_versions` | Version history with dates, authors, sizes |
| `drive_delta` | Incremental change tracking with delta tokens |
| `drive_thumbnail` | Thumbnail URLs (small/medium/large) for images and documents |

### SharePoint (3)

| Tool | Description |
|------|-------------|
| `drive_sites` | List or search SharePoint sites |
| `drive_site_libraries` | Document libraries for a SharePoint site |
| `drive_site_list` | Files in a SharePoint document library |

### Write (4, gated)

| Tool | Description |
|------|-------------|
| `drive_upload` | Upload a file (up to 4MB, base64 content) |
| `drive_create_folder` | Create a new folder |
| `drive_move` | Move and/or rename a file or folder |
| `drive_delete` | Delete to recycle bin (requires `confirm=true`) |

### Sharing (2)

| Tool | Description |
|------|-------------|
| `drive_share` | Create sharing links (view/edit/embed, org/anonymous scope) |
| `drive_permissions` | View who has access and existing sharing links |

## Delta Sync

Track file changes without re-listing entire directories:

```
# First call — returns current state + delta_link
drive_delta(folder_path="/Documents")

# Subsequent calls — returns only changes since last call
drive_delta(delta_link="aHR0cHM6Ly9ncmFwaC5taW...")
```

Returns created, modified, and deleted items with a fresh delta_link for the next call.

## Authentication

### Device Code (interactive)

```bash
export ONEDRIVE_AUTH_MODE="device_code"
export ONEDRIVE_TENANT_ID="your-tenant-id"
export ONEDRIVE_CLIENT_ID="your-client-id"
```

### Client Credentials (headless)

```bash
export ONEDRIVE_AUTH_MODE="client_credentials"
export ONEDRIVE_TENANT_ID="your-tenant-id"
export ONEDRIVE_CLIENT_ID="your-client-id"
export ONEDRIVE_CLIENT_SECRET="your-client-secret"
```

## Security Model

| Layer | Behaviour |
|-------|-----------|
| **Write gate** | All mutations disabled by default (`ONEDRIVE_WRITE_ENABLED=true`) |
| **Delete safety** | `drive_delete` requires `confirm=true` (moves to recycle bin) |
| **Credential scrubbing** | Tokens never appear in tool output or error messages |
| **Bearer auth** | Optional `ONEDRIVE_MCP_API_TOKEN` for HTTP transport |

## Claude Desktop Config

```json
{
  "mcpServers": {
    "onedrive": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/onedrive-blade-mcp", "onedrive-blade-mcp"],
      "env": {
        "ONEDRIVE_TENANT_ID": "your-tenant-id",
        "ONEDRIVE_CLIENT_ID": "your-client-id",
        "ONEDRIVE_AUTH_MODE": "device_code"
      }
    }
  }
}
```

## Claude Code

```bash
claude mcp add onedrive -- uv run --directory /path/to/onedrive-blade-mcp onedrive-blade-mcp
```

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `ONEDRIVE_TENANT_ID` | Yes | — | Azure AD tenant ID |
| `ONEDRIVE_CLIENT_ID` | Yes | — | App registration client ID |
| `ONEDRIVE_CLIENT_SECRET` | For client_credentials | — | App registration client secret |
| `ONEDRIVE_AUTH_MODE` | No | `device_code` | `device_code` or `client_credentials` |
| `ONEDRIVE_WRITE_ENABLED` | No | `false` | Enable write/delete tools |
| `ONEDRIVE_MCP_TRANSPORT` | No | `stdio` | `stdio` or `http` |
| `ONEDRIVE_MCP_HOST` | No | `127.0.0.1` | HTTP bind address |
| `ONEDRIVE_MCP_PORT` | No | `8002` | HTTP port |
| `ONEDRIVE_MCP_API_TOKEN` | No | — | Bearer token for HTTP auth |

## Architecture

```
src/onedrive_blade_mcp/
├── server.py       — FastMCP 2.0 server, 18 @mcp.tool decorators
├── client.py       — Graph API client with dual auth, write gate, delta sync
├── formatters.py   — Token-efficient output (pipe-delimited, null omission)
├── models.py       — Config, constants
└── auth.py         — Device code + client credentials, bearer middleware
```

Built with [FastMCP 2.0](https://github.com/jlowin/fastmcp) and [httpx](https://github.com/encode/httpx).

## Development

```bash
uv sync               # Install dependencies
uv run onedrive-blade-mcp   # Run locally (stdio)
uv run ruff check .    # Lint
uv run pytest tests/   # Tests (mocked, no OneDrive needed)
```

## License

MIT
