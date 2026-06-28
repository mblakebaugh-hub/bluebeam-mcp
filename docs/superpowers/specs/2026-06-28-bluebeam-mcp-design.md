# Bluebeam Revu MCP Server — Design Spec

**Date:** 2026-06-28
**Status:** Approved

---

## Overview

A Python MCP server that exposes Bluebeam Revu 21 (desktop) functionality to Claude via Windows COM automation. Claude can open and manage PDF documents, read and write markups, control layers, and run workflow operations (flatten, export summary).

---

## Architecture

```
bluebeam-mcp/
├── server.py               # MCP server entry point, tool registration
├── com_thread.py           # Dedicated COM STA thread + dispatch queue
├── bluebeam_service.py     # Domain methods (documents, markups, layers, workflows)
├── tools/
│   ├── __init__.py
│   ├── documents.py        # open, close, save, list
│   ├── markups.py          # list, add (text box / callout / stamp), delete
│   ├── layers.py           # list, set visibility, add
│   └── workflows.py        # flatten, export markup summary CSV
├── tests/
│   ├── test_service.py     # Unit tests with mocked COM object
│   └── test_integration.py # Integration tests (require Revu 21 installed)
├── requirements.txt
└── README.md
```

### Layer responsibilities

| Layer | File | Responsibility |
|-------|------|----------------|
| MCP transport | `server.py` | Register tools, parse args, return results/errors to Claude |
| Domain logic | `bluebeam_service.py` | Translate domain intents to COM calls; owns reconnect logic |
| COM thread | `com_thread.py` | Single STA thread; all COM objects live here; queue-based dispatch |
| Tool modules | `tools/*.py` | Define MCP tool schemas and delegate to `BluebeamService` |

### COM threading model

Python's `pywin32` COM objects must be used on the thread that called `pythoncom.CoInitialize()`. The MCP framework may call tools from any thread, so:

1. At server startup, `com_thread.py` starts a daemon thread and calls `CoInitialize()` on it.
2. `BluebeamService` methods submit a `(callable, Future)` pair to a `queue.Queue`.
3. The COM thread loops on `queue.get()`, executes the callable against the live COM object, and resolves the Future.
4. The calling thread blocks on `Future.result(timeout=30)`.

### Connection strategy

On first use (lazy init on the COM thread):

1. Try `win32com.client.GetActiveObject("Revu.Application")` — attach to running instance.
2. On `COMError`, try `win32com.client.Dispatch("Revu.Application")` — launch Revu.
3. If both fail, raise `BluebeamNotAvailableError("Bluebeam Revu 21 is not installed or could not start")`.

On mid-session loss (COM object returns error on next call):

- Attempt one reconnect using the same strategy above.
- If reconnect fails, return: `"Lost connection to Revu — please reopen it"`.

---

## MCP Tools

All tool names are prefixed `bb_` to avoid collisions with other MCP servers.

### Documents

| Tool | Input | Output |
|------|-------|--------|
| `bb_open_document` | `path: str` | `{page_count: int}` |
| `bb_close_document` | `path: str` | `{success: bool}` |
| `bb_save_document` | `path: str \| null` (null = active doc) | `{success: bool}` |
| `bb_list_open_documents` | *(none)* | `[{path: str, page_count: int}]` |

### Markups

| Tool | Input | Output |
|------|-------|--------|
| `bb_list_markups` | `path: str`, `page: int \| null` (null = all pages) | `[MarkupRecord]` |
| `bb_add_text_box` | `path: str`, `page: int`, `x: float`, `y: float`, `width: float`, `height: float`, `text: str`, `author: str \| null` | `{markup_id: str}` |
| `bb_add_callout` | `path: str`, `page: int`, `x: float`, `y: float`, `text: str`, `author: str \| null` | `{markup_id: str}` |
| `bb_add_stamp` | `path: str`, `page: int`, `stamp_name: str`, `x: float`, `y: float` | `{markup_id: str}` |

> **Coordinates:** All `x`, `y`, `width`, `height` values are in **PDF points** (1/72 inch), measured from the **bottom-left corner** of the page — the standard PDF coordinate system. `stamp_name` must exactly match a stamp already present in Revu's stamp library; use `bb_list_markups` with type filter or check Revu's Stamps panel for available names.
| `bb_delete_markup` | `path: str`, `markup_id: str` | `{success: bool}` |

**MarkupRecord schema:**
```json
{
  "id": "string",
  "type": "string",
  "page": "integer",
  "author": "string",
  "date": "ISO8601 string",
  "subject": "string",
  "comment": "string",
  "x": "float",
  "y": "float"
}
```

### Layers

| Tool | Input | Output |
|------|-------|--------|
| `bb_list_layers` | `path: str` | `[{name: str, visible: bool}]` |
| `bb_set_layer_visibility` | `path: str`, `layer_name: str`, `visible: bool` | `{success: bool}` |
| `bb_add_layer` | `path: str`, `layer_name: str` | `{success: bool}` |

### Workflows

| Tool | Input | Output |
|------|-------|--------|
| `bb_flatten_document` | `path: str` | `{success: bool}` |
| `bb_export_markup_summary` | `path: str`, `output_path: str` | `{rows_written: int}` |

---

## Error Handling

Every tool follows this pattern:

```python
try:
    result = service.some_method(...)
    return result
except BluebeamNotAvailableError as e:
    raise McpError(str(e))
except BluebeamDocumentError as e:
    raise McpError(str(e))
except Exception as e:
    raise McpError(f"Unexpected error: {e}")
```

**Error types:**

| Exception | When raised |
|-----------|-------------|
| `BluebeamNotAvailableError` | Revu not installed / couldn't launch / lost connection |
| `BluebeamDocumentError` | File not found, page out of range, markup ID not found, layer not found |

Arguments are validated before the COM call (e.g., file path exists, page number > 0) to produce cleaner errors than raw COM exceptions.

---

## Dependencies

```
mcp>=1.0.0
pywin32>=306
pytest
pytest-mock
```

Requires **Bluebeam Revu 21** installed on Windows.

---

## Testing

**Unit tests** (`tests/test_service.py`):
- `BluebeamService` accepts an optional `_app` constructor argument.
- In tests, pass a `MagicMock()` as `_app` to simulate COM responses without Revu installed.
- Covers: connection fallback logic, each domain method's COM call pattern, error mapping.

**Integration tests** (`tests/test_integration.py`):
- Tagged `@pytest.mark.integration` — skipped unless `--integration` flag passed to pytest.
- Require Revu 21 installed and a `tests/fixtures/sample.pdf` file.
- Cover: full open → markup → list → delete → close → flatten round-trip.

---

## Out of Scope (v1)

- Bluebeam Studio (cloud) sessions
- PDF rendering / page images
- Digital signatures
- Form fields
- Batch processing multiple files simultaneously
