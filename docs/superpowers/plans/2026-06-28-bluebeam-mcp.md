# Bluebeam Revu MCP Server — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Python MCP server that exposes Bluebeam Revu 21 (desktop) to Claude via Windows COM automation, covering document management, markups, layers, and workflow operations.

**Architecture:** A `COMThread` singleton owns all COM objects on a dedicated STA thread; `BluebeamService` dispatches domain calls to that thread via a queue; MCP tool modules register FastMCP tools that delegate to `BluebeamService`.

**Tech Stack:** Python 3.11+, `mcp>=1.0.0` (FastMCP), `pywin32>=306`, `pytest>=7.0`, `pytest-mock>=3.0`

## Global Constraints

- Windows 11 only — COM automation requires Windows
- Bluebeam Revu 21 (subscription) must be installed; ProgID: `"Revu.Application"`
- All COM objects must be created and used exclusively on the `COMThread` STA thread
- All MCP tool names prefixed `bb_`
- Markup coordinates in PDF points (1/72 inch), origin bottom-left of page
- `BluebeamService` must accept `_app` and `_com` constructor overrides for unit testing
- Integration tests tagged `@pytest.mark.integration`, skipped by default via `addopts`
- **COM method names throughout this plan are best-effort estimates based on common Bluebeam API patterns. Before implementing Tasks 4–7, run `python -m win32com.client.makepy "Revu.Application"` to generate stubs from the real type library and verify each method name.**

---

### Task 1: Project Scaffold

**Files:**
- Create: `requirements.txt`
- Create: `exceptions.py`
- Create: `pytest.ini`
- Create: `tests/__init__.py` (empty)
- Create: `tests/conftest.py`
- Create: `tools/__init__.py` (empty)

**Interfaces:**
- Produces: `BluebeamNotAvailableError`, `BluebeamDocumentError` — imported by all subsequent tasks
- Produces: `SyncDispatcher`, `mock_app` fixture, `service` fixture — used in all unit tests

- [ ] **Step 1: Create requirements.txt**

```
mcp>=1.0.0
pywin32>=306
pytest>=7.0
pytest-mock>=3.0
```

- [ ] **Step 2: Create exceptions.py**

```python
class BluebeamNotAvailableError(Exception):
    """Revu is not installed, not running, or connection was lost."""


class BluebeamDocumentError(Exception):
    """File not found, page out of range, markup/layer not found."""
```

- [ ] **Step 3: Create pytest.ini**

```ini
[pytest]
markers =
    integration: requires Bluebeam Revu 21 installed
addopts = -m "not integration"
```

- [ ] **Step 4: Create tests/conftest.py**

```python
import pytest
from unittest.mock import MagicMock


class SyncDispatcher:
    """Runs COM-thread callables synchronously for unit testing."""

    def run(self, fn, timeout=30):
        return fn()


@pytest.fixture
def mock_app():
    return MagicMock()


@pytest.fixture
def service(mock_app):
    from bluebeam_service import BluebeamService
    return BluebeamService(_app=mock_app, _com=SyncDispatcher())
```

- [ ] **Step 5: Create empty files**

Create `tests/__init__.py` and `tools/__init__.py` — both empty.

- [ ] **Step 6: Install dependencies**

```
pip install -r requirements.txt
```

Expected: all packages install without error.

- [ ] **Step 7: Verify pytest collects nothing yet**

```
pytest --collect-only
```

Expected: `no tests ran`

- [ ] **Step 8: Commit**

```bash
git add requirements.txt exceptions.py pytest.ini tests/ tools/
git commit -m "feat: scaffold project structure, exceptions, and test config"
```

---

### Task 2: COM Thread Dispatcher

**Files:**
- Create: `com_thread.py`
- Create: `tests/test_com_thread.py`

**Interfaces:**
- Produces: `COMThread` class with `.run(fn, timeout=30) -> Any` — used by `BluebeamService` in Task 3

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_com_thread.py
import threading
import pytest
from com_thread import COMThread


def test_run_executes_callable_on_background_thread():
    com = COMThread()
    caller_thread = threading.current_thread()
    captured = {}

    def fn():
        captured["thread"] = threading.current_thread()
        return 42

    result = com.run(fn)
    assert result == 42
    assert captured["thread"] is not caller_thread


def test_run_propagates_exceptions():
    com = COMThread()

    def fn():
        raise ValueError("boom")

    with pytest.raises(ValueError, match="boom"):
        com.run(fn)


def test_run_timeout_raises():
    import time
    com = COMThread()

    def fn():
        time.sleep(5)
        return 1

    with pytest.raises(Exception):
        com.run(fn, timeout=0.1)
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_com_thread.py -v
```

Expected: `ImportError: No module named 'com_thread'`

- [ ] **Step 3: Implement com_thread.py**

```python
import queue
import threading
from concurrent.futures import Future

import pythoncom


class COMThread:
    """Dedicated STA thread for all COM interactions."""

    def __init__(self):
        self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self):
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, future = item
                try:
                    future.set_result(fn())
                except Exception as exc:
                    future.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()

    def run(self, fn, timeout=30):
        """Submit fn to the COM thread and block until it returns."""
        future = Future()
        self._queue.put((fn, future))
        return future.result(timeout=timeout)

    def stop(self):
        self._queue.put(None)
        self._thread.join()
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_com_thread.py -v
```

Expected: 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add com_thread.py tests/test_com_thread.py
git commit -m "feat: add COMThread STA dispatcher"
```

---

### Task 3: BluebeamService — Connection Layer

**Files:**
- Create: `bluebeam_service.py`
- Create: `tests/test_service.py`

**Interfaces:**
- Consumes: `COMThread.run(fn)` from Task 2; `BluebeamNotAvailableError` from Task 1
- Produces: `BluebeamService(_app=None, _com=None)` with internal `_get_or_connect()` and `_call(fn)` — used by Tasks 4–7

**Before writing any COM calls**, generate Bluebeam's type stubs to verify method names:

```
python -m win32com.client.makepy "Revu.Application"
```

This prints the path to a generated `.py` file listing all COM methods and properties. Open it and cross-reference every method name used in Tasks 4–7.

- [ ] **Step 1: Write the failing connection tests**

```python
# tests/test_service.py
import pytest
from unittest.mock import MagicMock, patch
from exceptions import BluebeamNotAvailableError
from tests.conftest import SyncDispatcher


def test_uses_injected_app(service, mock_app):
    app = service._call(lambda a: a)
    assert app is mock_app


def test_connect_attaches_to_running_revu():
    from bluebeam_service import BluebeamService
    mock_revu = MagicMock()
    with patch("bluebeam_service.win32com.client.GetActiveObject", return_value=mock_revu):
        svc = BluebeamService(_com=SyncDispatcher())
        app = svc._call(lambda a: a)
    assert app is mock_revu


def test_connect_launches_revu_when_not_running():
    import pywintypes
    from bluebeam_service import BluebeamService
    mock_revu = MagicMock()
    with patch("bluebeam_service.win32com.client.GetActiveObject",
               side_effect=pywintypes.com_error()), \
         patch("bluebeam_service.win32com.client.Dispatch", return_value=mock_revu):
        svc = BluebeamService(_com=SyncDispatcher())
        app = svc._call(lambda a: a)
    assert app is mock_revu


def test_raises_when_revu_unavailable():
    import pywintypes
    from bluebeam_service import BluebeamService
    with patch("bluebeam_service.win32com.client.GetActiveObject",
               side_effect=pywintypes.com_error()), \
         patch("bluebeam_service.win32com.client.Dispatch",
               side_effect=pywintypes.com_error()):
        svc = BluebeamService(_com=SyncDispatcher())
        with pytest.raises(BluebeamNotAvailableError):
            svc._call(lambda a: a)
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_service.py -v
```

Expected: `ImportError: No module named 'bluebeam_service'`

- [ ] **Step 3: Implement bluebeam_service.py (connection only)**

```python
import pywintypes
import win32com.client

from com_thread import COMThread
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


class BluebeamService:
    def __init__(self, _app=None, _com=None):
        self._test_app = _app
        self._com = _com if _com is not None else COMThread()
        self._app = None  # COM object — only touched inside COM thread lambdas

    def _connect(self):
        """Attach to running Revu or launch it. Call only from COM thread."""
        try:
            return win32com.client.GetActiveObject("Revu.Application")
        except pywintypes.com_error:
            try:
                return win32com.client.Dispatch("Revu.Application")
            except pywintypes.com_error:
                raise BluebeamNotAvailableError(
                    "Bluebeam Revu 21 is not installed or could not start"
                )

    def _get_or_connect(self):
        """Return cached or fresh COM app. Call only from COM thread."""
        if self._test_app is not None:
            return self._test_app
        if self._app is None:
            self._app = self._connect()
        return self._app

    def _call(self, fn):
        """Dispatch fn(app) to COM thread; retry once on COM error."""
        def _do():
            try:
                return fn(self._get_or_connect())
            except pywintypes.com_error:
                self._app = None
                try:
                    return fn(self._get_or_connect())
                except pywintypes.com_error:
                    raise BluebeamNotAvailableError(
                        "Lost connection to Revu — please reopen it"
                    )
        return self._com.run(_do)
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_service.py -v
```

Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add bluebeam_service.py tests/test_service.py
git commit -m "feat: add BluebeamService with COM connection and reconnect logic"
```

---

### Task 4: BluebeamService — Document Methods

**Files:**
- Modify: `bluebeam_service.py`
- Modify: `tests/test_service.py`

**Interfaces:**
- Consumes: `BluebeamService._call(fn)` from Task 3
- Produces:
  - `open_document(path: str) -> dict` — `{"page_count": int}`
  - `close_document(path: str) -> dict` — `{"success": bool}`
  - `save_document(path: str | None) -> dict` — `{"success": bool}`
  - `list_open_documents() -> list[dict]` — `[{"path": str, "page_count": int}]`

**COM NOTE — verify these names against the generated stubs:**
- `app.Open(path)` — opens a PDF; returns the document object
- `app.Close(path)` — closes a document by file path
- `app.Save(path)` — saves a document; `None` saves the active document
- `app.Documents` — iterable of open document objects
- `doc.FilePath` — str, the document's absolute file path
- `doc.PageCount` — int, number of pages

- [ ] **Step 1: Write the failing document tests**

Append to `tests/test_service.py`:

```python
import os
from unittest.mock import MagicMock
from exceptions import BluebeamDocumentError


def test_open_document_raises_if_file_missing(service):
    with pytest.raises(BluebeamDocumentError, match="File not found"):
        service.open_document("C:\\does_not_exist.pdf")


def test_open_document_calls_com_open(service, mock_app, tmp_path):
    pdf = tmp_path / "test.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    mock_doc = MagicMock()
    mock_doc.PageCount = 5
    mock_app.Open.return_value = mock_doc

    result = service.open_document(str(pdf))

    mock_app.Open.assert_called_once_with(str(pdf))
    assert result == {"page_count": 5}


def test_close_document(service, mock_app):
    result = service.close_document("C:\\some.pdf")
    mock_app.Close.assert_called_once_with("C:\\some.pdf")
    assert result == {"success": True}


def test_save_document_with_path(service, mock_app):
    result = service.save_document("C:\\some.pdf")
    mock_app.Save.assert_called_once_with("C:\\some.pdf")
    assert result == {"success": True}


def test_save_document_active(service, mock_app):
    result = service.save_document(None)
    mock_app.Save.assert_called_once_with(None)
    assert result == {"success": True}


def test_list_open_documents(service, mock_app):
    doc1 = MagicMock(FilePath="C:\\a.pdf", PageCount=3)
    doc2 = MagicMock(FilePath="C:\\b.pdf", PageCount=7)
    mock_app.Documents = [doc1, doc2]

    result = service.list_open_documents()

    assert result == [
        {"path": "C:\\a.pdf", "page_count": 3},
        {"path": "C:\\b.pdf", "page_count": 7},
    ]
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_service.py -v -k "document"
```

Expected: FAIL — methods not defined on `BluebeamService`

- [ ] **Step 3: Add document methods to bluebeam_service.py**

Add `import os` at the top. Then add these methods to the `BluebeamService` class:

```python
def open_document(self, path: str) -> dict:
    if not os.path.exists(path):
        raise BluebeamDocumentError(f"File not found: {path}")
    return self._call(lambda app: {"page_count": app.Open(path).PageCount})

def close_document(self, path: str) -> dict:
    self._call(lambda app: app.Close(path))
    return {"success": True}

def save_document(self, path) -> dict:
    self._call(lambda app: app.Save(path))
    return {"success": True}

def list_open_documents(self) -> list:
    return self._call(
        lambda app: [
            {"path": doc.FilePath, "page_count": doc.PageCount}
            for doc in app.Documents
        ]
    )
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_service.py -v -k "document"
```

Expected: 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add bluebeam_service.py tests/test_service.py
git commit -m "feat: add document methods to BluebeamService"
```

---

### Task 5: BluebeamService — Markup Methods

**Files:**
- Modify: `bluebeam_service.py`
- Modify: `tests/test_service.py`

**Interfaces:**
- Consumes: `BluebeamService._call(fn)` from Task 3
- Produces:
  - `list_markups(path: str, page: int | None) -> list[dict]`
  - `add_text_box(path, page, x, y, width, height, text, author) -> dict` — `{"markup_id": str}`
  - `add_callout(path, page, x, y, text, author) -> dict` — `{"markup_id": str}`
  - `add_stamp(path, page, stamp_name, x, y) -> dict` — `{"markup_id": str}`
  - `delete_markup(path, markup_id) -> dict` — `{"success": bool}`

**COM NOTE — verify these names against the generated stubs:**
- `app.GetDocument(path)` — returns document object by file path
- `doc.Markups` — iterable of markup objects
- `markup.ID` — str, unique identifier
- `markup.Type` — str, markup type name
- `markup.PageNumber` — int, 1-based page number
- `markup.Author` — str
- `markup.Subject` — str
- `markup.Comments` — str
- `markup.Date` — str (ISO 8601 or similar)
- `markup.Rect.X`, `markup.Rect.Y` — float, bottom-left position in PDF points
- `doc.AddTextBox(page, x, y, width, height, text, author)` — returns markup ID str
- `doc.AddCallout(page, x, y, text, author)` — returns markup ID str
- `doc.AddStamp(page, stamp_name, x, y)` — returns markup ID str
- `doc.DeleteMarkup(markup_id)` — deletes markup by ID

- [ ] **Step 1: Write the failing markup tests**

Append to `tests/test_service.py`:

```python
def _make_markup(id="m1", type="TextBox", page=1, author="Alice",
                 subject="", comments="Hello", date="2026-06-28",
                 x=10.0, y=20.0):
    m = MagicMock()
    m.ID = id
    m.Type = type
    m.PageNumber = page
    m.Author = author
    m.Subject = subject
    m.Comments = comments
    m.Date = date
    m.Rect = MagicMock(X=x, Y=y)
    return m


def test_list_markups_all_pages(service, mock_app):
    doc = MagicMock()
    doc.Markups = [_make_markup("m1", page=1), _make_markup("m2", page=2)]
    mock_app.GetDocument.return_value = doc

    result = service.list_markups("C:\\a.pdf", page=None)

    assert len(result) == 2
    assert result[0]["id"] == "m1"
    assert result[1]["id"] == "m2"


def test_list_markups_filter_page(service, mock_app):
    doc = MagicMock()
    doc.Markups = [_make_markup("m1", page=1), _make_markup("m2", page=2)]
    mock_app.GetDocument.return_value = doc

    result = service.list_markups("C:\\a.pdf", page=1)

    assert len(result) == 1
    assert result[0]["id"] == "m1"


def test_add_text_box(service, mock_app):
    doc = MagicMock()
    doc.AddTextBox.return_value = "markup-123"
    mock_app.GetDocument.return_value = doc

    result = service.add_text_box("C:\\a.pdf", 1, 10.0, 20.0, 100.0, 50.0, "Note", "Alice")

    doc.AddTextBox.assert_called_once_with(1, 10.0, 20.0, 100.0, 50.0, "Note", "Alice")
    assert result == {"markup_id": "markup-123"}


def test_add_callout(service, mock_app):
    doc = MagicMock()
    doc.AddCallout.return_value = "markup-456"
    mock_app.GetDocument.return_value = doc

    result = service.add_callout("C:\\a.pdf", 1, 10.0, 20.0, "See this", None)

    doc.AddCallout.assert_called_once_with(1, 10.0, 20.0, "See this", None)
    assert result == {"markup_id": "markup-456"}


def test_add_stamp(service, mock_app):
    doc = MagicMock()
    doc.AddStamp.return_value = "markup-789"
    mock_app.GetDocument.return_value = doc

    result = service.add_stamp("C:\\a.pdf", 1, "APPROVED", 50.0, 60.0)

    doc.AddStamp.assert_called_once_with(1, "APPROVED", 50.0, 60.0)
    assert result == {"markup_id": "markup-789"}


def test_delete_markup(service, mock_app):
    doc = MagicMock()
    mock_app.GetDocument.return_value = doc

    result = service.delete_markup("C:\\a.pdf", "markup-123")

    doc.DeleteMarkup.assert_called_once_with("markup-123")
    assert result == {"success": True}
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_service.py -v -k "markup"
```

Expected: FAIL — methods not defined

- [ ] **Step 3: Add markup methods to bluebeam_service.py**

```python
def list_markups(self, path: str, page) -> list:
    def _do(app):
        doc = app.GetDocument(path)
        return [
            {
                "id": m.ID,
                "type": m.Type,
                "page": m.PageNumber,
                "author": m.Author,
                "subject": m.Subject,
                "comment": m.Comments,
                "date": m.Date,
                "x": m.Rect.X,
                "y": m.Rect.Y,
            }
            for m in doc.Markups
            if page is None or m.PageNumber == page
        ]
    return self._call(_do)

def add_text_box(self, path: str, page: int, x: float, y: float,
                 width: float, height: float, text: str, author) -> dict:
    def _do(app):
        doc = app.GetDocument(path)
        markup_id = doc.AddTextBox(page, x, y, width, height, text, author)
        return {"markup_id": markup_id}
    return self._call(_do)

def add_callout(self, path: str, page: int, x: float, y: float,
                text: str, author) -> dict:
    def _do(app):
        doc = app.GetDocument(path)
        markup_id = doc.AddCallout(page, x, y, text, author)
        return {"markup_id": markup_id}
    return self._call(_do)

def add_stamp(self, path: str, page: int, stamp_name: str,
              x: float, y: float) -> dict:
    def _do(app):
        doc = app.GetDocument(path)
        markup_id = doc.AddStamp(page, stamp_name, x, y)
        return {"markup_id": markup_id}
    return self._call(_do)

def delete_markup(self, path: str, markup_id: str) -> dict:
    def _do(app):
        doc = app.GetDocument(path)
        doc.DeleteMarkup(markup_id)
        return {"success": True}
    return self._call(_do)
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_service.py -v -k "markup"
```

Expected: 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add bluebeam_service.py tests/test_service.py
git commit -m "feat: add markup methods to BluebeamService"
```

---

### Task 6: BluebeamService — Layer Methods

**Files:**
- Modify: `bluebeam_service.py`
- Modify: `tests/test_service.py`

**Interfaces:**
- Produces:
  - `list_layers(path: str) -> list[dict]` — `[{"name": str, "visible": bool}]`
  - `set_layer_visibility(path: str, layer_name: str, visible: bool) -> dict` — `{"success": bool}`
  - `add_layer(path: str, layer_name: str) -> dict` — `{"success": bool}`

**COM NOTE — verify these names against the generated stubs:**
- `app.GetDocument(path)` — document object (same as Task 5)
- `doc.Layers` — iterable of layer objects
- `layer.Name` — str
- `layer.Visible` — bool; setting `layer.Visible = True/False` changes visibility
- `doc.AddLayer(name)` — creates a new layer

- [ ] **Step 1: Write the failing layer tests**

Append to `tests/test_service.py`:

```python
def test_list_layers(service, mock_app):
    doc = MagicMock()
    doc.Layers = [MagicMock(Name="Electrical", Visible=True),
                  MagicMock(Name="Plumbing", Visible=False)]
    mock_app.GetDocument.return_value = doc

    result = service.list_layers("C:\\a.pdf")

    assert result == [
        {"name": "Electrical", "visible": True},
        {"name": "Plumbing", "visible": False},
    ]


def test_set_layer_visibility(service, mock_app):
    doc = MagicMock()
    layer = MagicMock(Name="Electrical", Visible=False)
    doc.Layers = [layer]
    mock_app.GetDocument.return_value = doc

    result = service.set_layer_visibility("C:\\a.pdf", "Electrical", True)

    assert layer.Visible is True
    assert result == {"success": True}


def test_set_layer_visibility_raises_if_not_found(service, mock_app):
    doc = MagicMock()
    doc.Layers = [MagicMock(Name="Electrical")]
    mock_app.GetDocument.return_value = doc

    with pytest.raises(BluebeamDocumentError, match="Layer not found: Plumbing"):
        service.set_layer_visibility("C:\\a.pdf", "Plumbing", True)


def test_add_layer(service, mock_app):
    doc = MagicMock()
    mock_app.GetDocument.return_value = doc

    result = service.add_layer("C:\\a.pdf", "New Layer")

    doc.AddLayer.assert_called_once_with("New Layer")
    assert result == {"success": True}
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_service.py -v -k "layer"
```

Expected: FAIL — methods not defined

- [ ] **Step 3: Add layer methods to bluebeam_service.py**

```python
def list_layers(self, path: str) -> list:
    return self._call(
        lambda app: [
            {"name": layer.Name, "visible": layer.Visible}
            for layer in app.GetDocument(path).Layers
        ]
    )

def set_layer_visibility(self, path: str, layer_name: str, visible: bool) -> dict:
    def _do(app):
        doc = app.GetDocument(path)
        for layer in doc.Layers:
            if layer.Name == layer_name:
                layer.Visible = visible
                return {"success": True}
        raise BluebeamDocumentError(f"Layer not found: {layer_name}")
    return self._call(_do)

def add_layer(self, path: str, layer_name: str) -> dict:
    def _do(app):
        app.GetDocument(path).AddLayer(layer_name)
        return {"success": True}
    return self._call(_do)
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_service.py -v -k "layer"
```

Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add bluebeam_service.py tests/test_service.py
git commit -m "feat: add layer methods to BluebeamService"
```

---

### Task 7: BluebeamService — Workflow Methods

**Files:**
- Modify: `bluebeam_service.py`
- Modify: `tests/test_service.py`

**Interfaces:**
- Produces:
  - `flatten_document(path: str) -> dict` — `{"success": bool}`
  - `export_markup_summary(path: str, output_path: str) -> dict` — `{"rows_written": int}`

**COM NOTE — verify these names against the generated stubs:**
- `doc.Flatten()` — flattens all markups into the PDF content
- `doc.ExportMarkupSummary(output_path)` — exports a CSV; returns row count as int

- [ ] **Step 1: Write the failing workflow tests**

Append to `tests/test_service.py`:

```python
def test_flatten_document(service, mock_app):
    doc = MagicMock()
    mock_app.GetDocument.return_value = doc

    result = service.flatten_document("C:\\a.pdf")

    doc.Flatten.assert_called_once()
    assert result == {"success": True}


def test_export_markup_summary(service, mock_app):
    doc = MagicMock()
    doc.ExportMarkupSummary.return_value = 12
    mock_app.GetDocument.return_value = doc

    result = service.export_markup_summary("C:\\a.pdf", "C:\\summary.csv")

    doc.ExportMarkupSummary.assert_called_once_with("C:\\summary.csv")
    assert result == {"rows_written": 12}
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_service.py -v -k "flatten or export"
```

Expected: FAIL — methods not defined

- [ ] **Step 3: Add workflow methods to bluebeam_service.py**

```python
def flatten_document(self, path: str) -> dict:
    def _do(app):
        app.GetDocument(path).Flatten()
        return {"success": True}
    return self._call(_do)

def export_markup_summary(self, path: str, output_path: str) -> dict:
    def _do(app):
        rows = app.GetDocument(path).ExportMarkupSummary(output_path)
        return {"rows_written": rows}
    return self._call(_do)
```

- [ ] **Step 4: Run full unit test suite**

```
pytest tests/test_com_thread.py tests/test_service.py -v
```

Expected: all tests PASS

- [ ] **Step 5: Commit**

```bash
git add bluebeam_service.py tests/test_service.py
git commit -m "feat: add workflow methods to BluebeamService"
```

---

### Task 8: MCP Tool Modules

**Files:**
- Create: `tools/documents.py`
- Create: `tools/markups.py`
- Create: `tools/layers.py`
- Create: `tools/workflows.py`

**Interfaces:**
- Consumes: all `BluebeamService` methods from Tasks 4–7; `BluebeamNotAvailableError`, `BluebeamDocumentError` from Task 1
- Produces: `register_document_tools(mcp, service)`, `register_markup_tools(mcp, service)`, `register_layer_tools(mcp, service)`, `register_workflow_tools(mcp, service)`

Tool modules are thin wrappers — no unit tests. Coverage comes from integration tests in Task 10.

- [ ] **Step 1: Create tools/documents.py**

```python
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_document_tools(mcp, service):
    @mcp.tool()
    def bb_open_document(path: str) -> dict:
        """Open a PDF in Bluebeam Revu. Returns page_count."""
        try:
            return service.open_document(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_close_document(path: str) -> dict:
        """Close an open PDF in Bluebeam Revu."""
        try:
            return service.close_document(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_save_document(path: str = None) -> dict:
        """Save a document. Omit path to save the active document."""
        try:
            return service.save_document(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_list_open_documents() -> list:
        """List all PDFs currently open in Bluebeam Revu."""
        try:
            return service.list_open_documents()
        except BluebeamNotAvailableError as e:
            raise Exception(str(e))
```

- [ ] **Step 2: Create tools/markups.py**

```python
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_markup_tools(mcp, service):
    @mcp.tool()
    def bb_list_markups(path: str, page: int = None) -> list:
        """List markups in a document. Omit page to list all pages.
        Returns list of {id, type, page, author, subject, comment, date, x, y}."""
        try:
            return service.list_markups(path, page)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_text_box(path: str, page: int, x: float, y: float,
                        width: float, height: float, text: str,
                        author: str = None) -> dict:
        """Add a text box markup. x, y, width, height in PDF points (1/72 in), origin bottom-left."""
        try:
            return service.add_text_box(path, page, x, y, width, height, text, author)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_callout(path: str, page: int, x: float, y: float,
                       text: str, author: str = None) -> dict:
        """Add a callout (leader line + text box). x, y in PDF points, origin bottom-left."""
        try:
            return service.add_callout(path, page, x, y, text, author)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_stamp(path: str, page: int, stamp_name: str,
                     x: float, y: float) -> dict:
        """Apply a named stamp. stamp_name must match a stamp in Revu's stamp library."""
        try:
            return service.add_stamp(path, page, stamp_name, x, y)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_delete_markup(path: str, markup_id: str) -> dict:
        """Delete a markup by its ID (get IDs from bb_list_markups)."""
        try:
            return service.delete_markup(path, markup_id)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
```

- [ ] **Step 3: Create tools/layers.py**

```python
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_layer_tools(mcp, service):
    @mcp.tool()
    def bb_list_layers(path: str) -> list:
        """List all layers and their visibility. Returns [{name, visible}]."""
        try:
            return service.list_layers(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_set_layer_visibility(path: str, layer_name: str, visible: bool) -> dict:
        """Show or hide a layer by name."""
        try:
            return service.set_layer_visibility(path, layer_name, visible)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_layer(path: str, layer_name: str) -> dict:
        """Create a new layer in a document."""
        try:
            return service.add_layer(path, layer_name)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
```

- [ ] **Step 4: Create tools/workflows.py**

```python
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_workflow_tools(mcp, service):
    @mcp.tool()
    def bb_flatten_document(path: str) -> dict:
        """Flatten all markups into the PDF content (irreversible)."""
        try:
            return service.flatten_document(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_export_markup_summary(path: str, output_path: str) -> dict:
        """Export a CSV markup summary report. Returns rows_written count."""
        try:
            return service.export_markup_summary(path, output_path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
```

- [ ] **Step 5: Verify imports are clean**

```
python -c "from tools.documents import register_document_tools; print('OK')"
python -c "from tools.markups import register_markup_tools; print('OK')"
python -c "from tools.layers import register_layer_tools; print('OK')"
python -c "from tools.workflows import register_workflow_tools; print('OK')"
```

Expected: each prints `OK`

- [ ] **Step 6: Commit**

```bash
git add tools/
git commit -m "feat: add MCP tool modules for all four domains"
```

---

### Task 9: MCP Server Entry Point

**Files:**
- Create: `server.py`

**Interfaces:**
- Consumes: all `register_*_tools` functions from Task 8; `BluebeamService` from Task 3
- Produces: runnable MCP server at `server.py`

- [ ] **Step 1: Create server.py**

```python
from mcp.server.fastmcp import FastMCP

from bluebeam_service import BluebeamService
from tools.documents import register_document_tools
from tools.layers import register_layer_tools
from tools.markups import register_markup_tools
from tools.workflows import register_workflow_tools

mcp = FastMCP("bluebeam-mcp")
service = BluebeamService()

register_document_tools(mcp, service)
register_markup_tools(mcp, service)
register_layer_tools(mcp, service)
register_workflow_tools(mcp, service)

if __name__ == "__main__":
    mcp.run()
```

- [ ] **Step 2: Verify server imports without error**

```
python -c "import server; print('Server OK')"
```

Expected: prints `Server OK` with no errors

- [ ] **Step 3: Commit**

```bash
git add server.py
git commit -m "feat: add MCP server entry point, wire all 14 tools"
```

---

### Task 10: Integration Tests + README

**Files:**
- Create: `tests/test_integration.py`
- Create: `tests/fixtures/.gitkeep` (empty — PDF must be added manually)
- Create: `README.md`

**Interfaces:**
- Consumes: full running stack — Revu 21 installed, `tests/fixtures/sample.pdf` present

- [ ] **Step 1: Create tests/test_integration.py**

```python
import os
import pytest

SAMPLE_PDF = os.path.join(os.path.dirname(__file__), "fixtures", "sample.pdf")


@pytest.mark.integration
class TestDocumentRoundTrip:
    def test_open_and_close(self):
        from bluebeam_service import BluebeamService
        svc = BluebeamService()
        result = svc.open_document(SAMPLE_PDF)
        assert "page_count" in result
        assert result["page_count"] > 0
        close = svc.close_document(SAMPLE_PDF)
        assert close["success"] is True

    def test_markup_roundtrip(self):
        from bluebeam_service import BluebeamService
        svc = BluebeamService()
        svc.open_document(SAMPLE_PDF)

        add = svc.add_text_box(SAMPLE_PDF, 1, 50.0, 50.0, 200.0, 50.0, "Integration test", "Claude")
        assert "markup_id" in add
        markup_id = add["markup_id"]

        markups = svc.list_markups(SAMPLE_PDF, page=1)
        assert markup_id in [m["id"] for m in markups]

        delete = svc.delete_markup(SAMPLE_PDF, markup_id)
        assert delete["success"] is True

        svc.close_document(SAMPLE_PDF)

    def test_layer_roundtrip(self):
        from bluebeam_service import BluebeamService
        svc = BluebeamService()
        svc.open_document(SAMPLE_PDF)

        svc.add_layer(SAMPLE_PDF, "TestLayer")
        layers = svc.list_layers(SAMPLE_PDF)
        assert "TestLayer" in [l["name"] for l in layers]

        svc.set_layer_visibility(SAMPLE_PDF, "TestLayer", False)
        layers = svc.list_layers(SAMPLE_PDF)
        layer = next(l for l in layers if l["name"] == "TestLayer")
        assert layer["visible"] is False

        svc.close_document(SAMPLE_PDF)
```

- [ ] **Step 2: Create tests/fixtures/.gitkeep**

Create an empty file at `tests/fixtures/.gitkeep`. Then manually copy any small PDF to `tests/fixtures/sample.pdf` (export one from Revu or use any existing PDF).

- [ ] **Step 3: Verify integration tests are skipped by default**

```
pytest tests/ -v
```

Expected: integration tests are deselected by `addopts = -m "not integration"` in `pytest.ini`; all unit tests PASS

- [ ] **Step 4: Create README.md**

```markdown
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
```

- [ ] **Step 5: Run full unit test suite one final time**

```
pytest tests/test_com_thread.py tests/test_service.py -v
```

Expected: all tests PASS

- [ ] **Step 6: Commit**

```bash
git add tests/test_integration.py tests/fixtures/.gitkeep README.md
git commit -m "feat: add integration tests and README"
```
