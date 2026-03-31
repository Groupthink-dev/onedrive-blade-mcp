---
name: onedrive-blade-mcp
tools:
  - drive_info
  - drive_list
  - drive_read
  - drive_metadata
  - drive_search
  - drive_download
  - drive_versions
  - drive_delta
  - drive_thumbnail
  - drive_sites
  - drive_site_libraries
  - drive_site_list
  - drive_upload
  - drive_create_folder
  - drive_move
  - drive_delete
  - drive_share
  - drive_permissions
permissions:
  read:
    - drive_info
    - drive_list
    - drive_read
    - drive_metadata
    - drive_search
    - drive_download
    - drive_versions
    - drive_delta
    - drive_thumbnail
    - drive_sites
    - drive_site_libraries
    - drive_site_list
    - drive_permissions
  write:
    - drive_upload
    - drive_create_folder
    - drive_move
    - drive_delete
    - drive_share
---

# OneDrive & SharePoint Blade MCP — Skill Guide

File operations via Microsoft Graph API. Token-efficient, write-gated, supports both personal OneDrive and SharePoint document libraries.

## Token Efficiency Rules

1. **Use `drive_list` before `drive_read`** — browse to find the right file, don't guess paths
2. **Use `drive_search` for discovery** — faster than recursive listing
3. **Use `drive_metadata` for file info** — don't read content just to check size/dates
4. **Use `drive_delta` for monitoring** — incremental sync instead of repeated full listings
5. **Use `drive_download` for large files** — get a URL instead of reading content inline
6. **Limit `max_bytes`** — when reading files, only request what you need
7. **Use folder_path over folder_id** when the path is known — avoids a metadata lookup
8. **Use `drive_thumbnail`** for image preview — don't read full image content

## Quick Start

```
# Check drive status
drive_info

# Browse files
drive_list                                    # root folder
drive_list folder_path="/Documents"           # specific folder
drive_list folder_path="/Documents" limit=50  # more results

# Find files
drive_search query="budget 2026"
drive_search query="*.pdf" limit=10

# Read a file
drive_metadata item_path="/Documents/notes.txt"    # check details first
drive_read item_path="/Documents/notes.txt"         # read content

# Upload a file (requires ONEDRIVE_WRITE_ENABLED=true)
drive_upload file_name="report.txt" content_base64="..." folder_path="/Documents"

# Share a file
drive_share item_id="..." link_type="view" scope="organization"
```

## Tool Reference

### Meta
| Tool | Purpose |
|------|---------|
| `drive_info` | Drive type, quota usage, owner |

### Read Tools (always available)
| Tool | Purpose | Key Params |
|------|---------|------------|
| `drive_list` | List folder contents | `folder_path`, `folder_id`, `limit` |
| `drive_read` | Read file content (text or base64) | `item_path`, `item_id`, `max_bytes` |
| `drive_metadata` | Detailed file/folder metadata | `item_path`, `item_id` |
| `drive_search` | Search by name or content | `query`, `limit` |
| `drive_download` | Get download URL | `item_path`, `item_id` |
| `drive_versions` | Version history | `item_id` |
| `drive_delta` | Incremental change tracking | `delta_link`, `folder_path` |
| `drive_thumbnail` | Thumbnail URLs (images, docs) | `item_id` |
| `drive_sites` | List/search SharePoint sites | `query`, `limit` |
| `drive_site_libraries` | List document libraries | `site_id` |
| `drive_site_list` | List files in SP library | `site_id`, `drive_id`, `folder_id` |
| `drive_permissions` | View sharing permissions | `item_id` |

### Write Tools (require ONEDRIVE_WRITE_ENABLED=true)
| Tool | Purpose | Key Params |
|------|---------|------------|
| `drive_upload` | Upload file (≤4MB) | `file_name`, `content_base64`, `folder_path`, `conflict` |
| `drive_create_folder` | Create folder | `name`, `parent_path` |
| `drive_move` | Move or rename item | `item_id`, `destination_folder_id`, `new_name` |
| `drive_delete` | Delete item (recycle bin) | `item_id`, `confirm=true` |
| `drive_share` | Create sharing link | `item_id`, `link_type`, `scope`, `expiration` |

## Workflows

### File Discovery & Reading
```
drive_search query="quarterly report"
→ drive_metadata item_id="..."
→ drive_read item_id="..." max_bytes=500000
```

### Incremental Monitoring
```
# First call — get baseline
drive_delta folder_path="/Shared Documents"
→ save the delta_link from output

# Subsequent calls — only changes
drive_delta delta_link="<saved_link>"
```

### SharePoint Document Library Access
```
drive_sites query="Marketing"
→ drive_site_libraries site_id="..."
→ drive_site_list site_id="..." drive_id="..."
→ drive_read item_id="..."
```

### Upload & Share
```
drive_upload file_name="report.pdf" content_base64="..." folder_path="/Reports"
→ drive_share item_id="..." link_type="view" scope="organization"
```

## Common Parameters

| Parameter | Values | Notes |
|-----------|--------|-------|
| `item_path` | `/Documents/file.txt` | Path from drive root |
| `item_id` | Graph item ID | From list/search results |
| `folder_path` | `/Documents` | Omit for root |
| `limit` | 1–200 | Default 20 |
| `max_bytes` | bytes | Default 1MB |
| `conflict` | rename/replace/fail | Default: rename |
| `link_type` | view/edit/embed | Default: view |
| `scope` | anonymous/organization | Default: organization |

## Security Notes

- Write operations disabled by default — set `ONEDRIVE_WRITE_ENABLED=true`
- Delete requires explicit `confirm=true` parameter
- Sharing defaults to `organization` scope (not anonymous)
- Token cache stored at `~/.onedrive-blade-mcp/token_cache.json` with 0600 permissions
- HTTP transport supports bearer token auth via `ONEDRIVE_MCP_API_TOKEN`
- File reads capped at `max_bytes` to prevent memory exhaustion
- Upload limited to 4MB (Graph API simple upload limit)
