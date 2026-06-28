# bluebeam-mcp

MCP server exposing Bluebeam Revu 21 to Claude via Windows COM automation.

## Requirements

- Windows 11
- Bluebeam Revu 21 (subscription) installed
- Python 3.11+

## Setup

    pip install -r requirements.txt

## Run

    python server.py

Add to Claude Code MCP config (`~/.claude/settings.json`):

    {
      "mcpServers": {
        "bluebeam": {
          "command": "python",
          "args": ["C:/path/to/bluebeam-mcp/server.py"]
        }
      }
    }

## Test

    pytest                    # unit tests only (no Revu required)
    pytest -m integration     # requires Revu 21 + tests/fixtures/sample.pdf

## Tools

| Tool | Description |
|------|-------------|
| bb_open_document | Open a PDF in Revu |
| bb_close_document | Close an open PDF |
| bb_save_document | Save a document |
| bb_list_open_documents | List all open PDFs |
| bb_list_markups | List markups (optionally filter by page) |
| bb_add_text_box | Add a text box markup |
| bb_add_callout | Add a callout markup |
| bb_add_stamp | Apply a stamp |
| bb_delete_markup | Delete a markup by ID |
| bb_list_layers | List layers and visibility |
| bb_set_layer_visibility | Show or hide a layer |
| bb_add_layer | Create a new layer |
| bb_flatten_document | Flatten markups into PDF content |
| bb_export_markup_summary | Export CSV markup summary |
