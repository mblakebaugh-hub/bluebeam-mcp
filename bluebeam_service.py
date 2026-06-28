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
