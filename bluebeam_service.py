import os
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

    def flatten_document(self, path: str) -> dict:
        def _do(app):
            app.GetDocument(path).Flatten()
            return {"success": True}
        return self._call(_do)

    def export_markup_summary(self, path: str, output_path: str) -> dict:
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.isdir(output_dir):
            raise BluebeamDocumentError(f"Output directory does not exist: {output_dir}")
        def _do(app):
            rows = app.GetDocument(path).ExportMarkupSummary(output_path)
            return {"rows_written": rows}
        return self._call(_do)
