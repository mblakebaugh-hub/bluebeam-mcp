import os
import pytest
from unittest.mock import MagicMock, patch
from exceptions import BluebeamNotAvailableError, BluebeamDocumentError
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
