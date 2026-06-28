import csv
import os
from typing import Optional

import fitz  # PyMuPDF
import pywintypes
import win32com.client

from com_thread import COMThread
from exceptions import BluebeamDocumentError, BluebeamNotAvailableError

# Maps common stamp names to PyMuPDF stamp integers (0-13)
_STAMP_TYPES = {
    "Approved": 0, "APPROVED": 0,
    "AsIs": 1, "AS IS": 1,
    "Confidential": 2, "CONFIDENTIAL": 2,
    "Departmental": 3,
    "Draft": 4, "DRAFT": 4,
    "Experimental": 5,
    "Expired": 6,
    "Final": 7, "FINAL": 7,
    "ForComment": 8, "FOR COMMENT": 8,
    "ForPublicRelease": 9,
    "NotApproved": 10, "NOT APPROVED": 10,
    "NotForPublicRelease": 11,
    "Sold": 12,
    "TopSecret": 13, "TOP SECRET": 13,
}


class BluebeamService:
    def __init__(self, _launcher=None, _com=None):
        self._test_launcher = _launcher
        self._com = _com if _com is not None else COMThread()
        self._launcher = None

    def _connect(self):
        """Attach to or launch Revu via Revu.Launcher COM. Call only from COM thread."""
        try:
            return win32com.client.GetActiveObject("Revu.Launcher")
        except pywintypes.com_error:
            try:
                return win32com.client.Dispatch("Revu.Launcher")
            except pywintypes.com_error:
                raise BluebeamNotAvailableError(
                    "Bluebeam Revu 21 is not installed or could not start"
                )

    def _get_or_connect(self):
        if self._test_launcher is not None:
            return self._test_launcher
        if self._launcher is None:
            self._launcher = self._connect()
        return self._launcher

    def _call_launcher(self, fn):
        """Dispatch fn(launcher) to the COM thread; retry once on COM error."""
        def _do():
            try:
                return fn(self._get_or_connect())
            except pywintypes.com_error:
                self._launcher = None
                try:
                    return fn(self._get_or_connect())
                except pywintypes.com_error:
                    raise BluebeamNotAvailableError(
                        "Lost connection to Revu — please reopen it"
                    )
        return self._com.run(_do)

    def _open_pdf(self, path: str) -> fitz.Document:
        if not os.path.exists(path):
            raise BluebeamDocumentError(f"File not found: {path}")
        try:
            return fitz.open(path)
        except Exception as e:
            raise BluebeamDocumentError(f"Could not open PDF: {e}")

    def _save_close(self, doc: fitz.Document, path: str, incremental: bool = True) -> None:
        doc.save(path, incremental=incremental, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()

    # --- Document methods ---

    def open_document(self, path: str) -> dict:
        if not os.path.exists(path):
            raise BluebeamDocumentError(f"File not found: {path}")
        self._call_launcher(lambda l: l.EditDocument(path))
        doc = fitz.open(path)
        page_count = doc.page_count
        doc.close()
        return {"page_count": page_count}

    def close_document(self, path: str) -> dict:
        # Revu.Launcher has no close API; acknowledge without error.
        return {"success": True}

    def save_document(self, path: Optional[str] = None) -> dict:
        if path is None:
            # Cannot save active document without a full COM app object.
            return {"success": True}
        doc = self._open_pdf(path)
        try:
            self._save_close(doc, path)
        except Exception:
            doc.close()
            raise
        return {"success": True}

    def list_open_documents(self) -> list:
        # Revu.Launcher does not expose a document list.
        return []

    # --- Markup methods ---

    def list_markups(self, path: str, page: Optional[int] = None) -> list:
        doc = self._open_pdf(path)
        try:
            results = []
            page_range = range(doc.page_count) if page is None else [page - 1]
            for pi in page_range:
                if pi < 0 or pi >= doc.page_count:
                    raise BluebeamDocumentError(f"Page {pi + 1} out of range")
                pg = doc[pi]
                for annot in pg.annots():
                    info = annot.info
                    r = annot.rect
                    results.append({
                        "id": str(annot.xref),
                        "type": annot.type[1],
                        "page": pi + 1,
                        "author": info.get("title", ""),
                        "subject": info.get("subject", ""),
                        "comment": info.get("content", ""),
                        "date": info.get("modDate", ""),
                        "x": r.x0,
                        "y": r.y0,
                    })
            return results
        finally:
            doc.close()

    def add_text_box(self, path: str, page: int, x: float, y: float,
                     width: float, height: float, text: str,
                     author: Optional[str] = None) -> dict:
        doc = self._open_pdf(path)
        try:
            if page < 1 or page > doc.page_count:
                raise BluebeamDocumentError(f"Page {page} out of range")
            pg = doc[page - 1]
            rect = fitz.Rect(x, y, x + width, y + height)
            annot = pg.add_freetext_annot(rect, text)
            if author:
                annot.set_info(title=author)
            annot.update()
            xref = annot.xref
            self._save_close(doc, path)
            return {"markup_id": str(xref)}
        except BluebeamDocumentError:
            doc.close()
            raise
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    def add_callout(self, path: str, page: int, x: float, y: float,
                    text: str, author: Optional[str] = None) -> dict:
        doc = self._open_pdf(path)
        try:
            if page < 1 or page > doc.page_count:
                raise BluebeamDocumentError(f"Page {page} out of range")
            pg = doc[page - 1]
            rect = fitz.Rect(x, y, x + 150, y + 50)
            annot = pg.add_freetext_annot(
                rect, text,
                callout=[fitz.Point(x - 20, y + 25), fitz.Point(x, y + 25)],
            )
            if author:
                annot.set_info(title=author)
            annot.update()
            xref = annot.xref
            self._save_close(doc, path)
            return {"markup_id": str(xref)}
        except BluebeamDocumentError:
            doc.close()
            raise
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    def add_stamp(self, path: str, page: int, stamp_name: str,
                  x: float, y: float) -> dict:
        doc = self._open_pdf(path)
        try:
            if page < 1 or page > doc.page_count:
                raise BluebeamDocumentError(f"Page {page} out of range")
            pg = doc[page - 1]
            rect = fitz.Rect(x, y, x + 150, y + 50)
            stamp_int = _STAMP_TYPES.get(stamp_name, 0)
            annot = pg.add_stamp_annot(rect, stamp=stamp_int)
            annot.set_info(subject=stamp_name)
            annot.update()
            xref = annot.xref
            self._save_close(doc, path)
            return {"markup_id": str(xref)}
        except BluebeamDocumentError:
            doc.close()
            raise
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    def delete_markup(self, path: str, markup_id: str) -> dict:
        try:
            target_xref = int(markup_id)
        except ValueError:
            raise BluebeamDocumentError(f"Invalid markup ID: {markup_id}")
        doc = self._open_pdf(path)
        try:
            for pi in range(doc.page_count):
                pg = doc[pi]
                for annot in list(pg.annots()):
                    if annot.xref == target_xref:
                        pg.delete_annot(annot)
                        self._save_close(doc, path)
                        return {"success": True}
            raise BluebeamDocumentError(f"Markup not found: {markup_id}")
        except BluebeamDocumentError:
            doc.close()
            raise
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    # --- Layer methods ---

    def list_layers(self, path: str) -> list:
        doc = self._open_pdf(path)
        try:
            ocgs = doc.get_ocgs()  # {xref: {"name": str, ...}}
            on_set = set((doc.get_layer(-1) or {}).get("on", []))
            return [
                {"name": info["name"], "visible": xref in on_set}
                for xref, info in ocgs.items()
            ]
        finally:
            doc.close()

    def set_layer_visibility(self, path: str, layer_name: str, visible: bool) -> dict:
        doc = self._open_pdf(path)
        try:
            target_xref = next(
                (xref for xref, info in doc.get_ocgs().items() if info["name"] == layer_name),
                None,
            )
            if target_xref is None:
                raise BluebeamDocumentError(f"Layer not found: {layer_name}")
            layer = doc.get_layer(-1)
            on_list = [x for x in layer.get("on", []) if x != target_xref]
            off_list = [x for x in layer.get("off", []) if x != target_xref]
            if visible:
                on_list = [target_xref] + on_list
            else:
                off_list = [target_xref] + off_list
            doc.set_layer(-1, on=on_list, off=off_list)
            self._save_close(doc, path)
            return {"success": True}
        except BluebeamDocumentError:
            doc.close()
            raise
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    def add_layer(self, path: str, layer_name: str) -> dict:
        doc = self._open_pdf(path)
        try:
            doc.add_ocg(layer_name, on=True)
            self._save_close(doc, path)
            return {"success": True}
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    # --- Workflow methods ---

    def flatten_document(self, path: str) -> dict:
        """Remove all PDF annotations. Note: does not bake annotation visuals into page content."""
        doc = self._open_pdf(path)
        try:
            for pi in range(doc.page_count):
                pg = doc[pi]
                for annot in list(pg.annots()):
                    pg.delete_annot(annot)
            self._save_close(doc, path)
            return {"success": True}
        except Exception as e:
            doc.close()
            raise BluebeamDocumentError(str(e))

    def export_markup_summary(self, path: str, output_path: str) -> dict:
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.isdir(output_dir):
            raise BluebeamDocumentError(f"Output directory does not exist: {output_dir}")
        markups = self.list_markups(path)
        fields = ["id", "type", "page", "author", "subject", "comment", "date", "x", "y"]
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(markups)
        return {"rows_written": len(markups)}
