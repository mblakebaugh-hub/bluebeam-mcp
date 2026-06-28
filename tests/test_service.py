import os
import pytest
from unittest.mock import MagicMock, patch
from exceptions import BluebeamNotAvailableError, BluebeamDocumentError
from tests.conftest import SyncDispatcher


# --- Launcher / connection tests ---

def test_uses_injected_launcher(service, mock_launcher):
    launcher = service._call_launcher(lambda l: l)
    assert launcher is mock_launcher


def test_connect_attaches_to_running_revu():
    from bluebeam_service import BluebeamService
    mock_revu = MagicMock()
    with patch("bluebeam_service.win32com.client.GetActiveObject", return_value=mock_revu):
        svc = BluebeamService(_com=SyncDispatcher())
        launcher = svc._call_launcher(lambda l: l)
    assert launcher is mock_revu


def test_connect_launches_revu_when_not_running():
    import pywintypes
    from bluebeam_service import BluebeamService
    mock_revu = MagicMock()
    with patch("bluebeam_service.win32com.client.GetActiveObject",
               side_effect=pywintypes.com_error()), \
         patch("bluebeam_service.win32com.client.Dispatch", return_value=mock_revu):
        svc = BluebeamService(_com=SyncDispatcher())
        launcher = svc._call_launcher(lambda l: l)
    assert launcher is mock_revu


def test_raises_when_revu_unavailable():
    import pywintypes
    from bluebeam_service import BluebeamService
    with patch("bluebeam_service.win32com.client.GetActiveObject",
               side_effect=pywintypes.com_error()), \
         patch("bluebeam_service.win32com.client.Dispatch",
               side_effect=pywintypes.com_error()):
        svc = BluebeamService(_com=SyncDispatcher())
        with pytest.raises(BluebeamNotAvailableError):
            svc._call_launcher(lambda l: l)


def test_call_launcher_retries_on_com_error_and_succeeds(mock_launcher):
    import pywintypes
    from bluebeam_service import BluebeamService
    call_count = {"n": 0}

    def flaky_fn(launcher):
        call_count["n"] += 1
        if call_count["n"] == 1:
            raise pywintypes.com_error()
        return "ok"

    svc = BluebeamService(_launcher=mock_launcher, _com=SyncDispatcher())
    result = svc._call_launcher(flaky_fn)
    assert result == "ok"
    assert call_count["n"] == 2


def test_call_launcher_raises_not_available_after_two_com_errors():
    import pywintypes
    from bluebeam_service import BluebeamService

    def always_fails(launcher):
        raise pywintypes.com_error()

    svc = BluebeamService(_launcher=MagicMock(), _com=SyncDispatcher())
    with pytest.raises(BluebeamNotAvailableError, match="Lost connection"):
        svc._call_launcher(always_fails)


# --- Document method tests ---

def test_open_document_raises_if_file_missing(service):
    with pytest.raises(BluebeamDocumentError, match="File not found"):
        service.open_document("C:\\does_not_exist.pdf")


def test_open_document_opens_in_revu_and_returns_page_count(service, mock_launcher, sample_pdf):
    result = service.open_document(sample_pdf)
    mock_launcher.EditDocument.assert_called_once_with(sample_pdf)
    assert result == {"page_count": 1}


def test_close_document_returns_success(service):
    result = service.close_document("C:\\any.pdf")
    assert result == {"success": True}


def test_save_document_with_path(service, sample_pdf):
    result = service.save_document(sample_pdf)
    assert result == {"success": True}


def test_save_document_none_returns_success(service):
    result = service.save_document(None)
    assert result == {"success": True}


def test_list_open_documents_returns_empty(service):
    assert service.list_open_documents() == []


# --- Markup method tests ---

def test_list_markups_empty(service, sample_pdf):
    result = service.list_markups(sample_pdf)
    assert result == []


def test_list_markups_all_pages(service, sample_pdf):
    service.add_text_box(sample_pdf, 1, 10.0, 10.0, 100.0, 30.0, "Page 1 note", None)
    result = service.list_markups(sample_pdf)
    assert len(result) == 1
    assert result[0]["page"] == 1


def test_list_markups_filter_page(service, tmp_path):
    import fitz
    path = str(tmp_path / "two_pages.pdf")
    doc = fitz.open()
    doc.new_page()
    doc.new_page()
    doc.save(path)
    doc.close()
    service.add_text_box(path, 1, 10.0, 10.0, 100.0, 30.0, "Page 1 note", None)
    assert service.list_markups(path, page=1) != []
    assert service.list_markups(path, page=2) == []


def test_add_text_box_returns_markup_id(service, sample_pdf):
    result = service.add_text_box(sample_pdf, 1, 50.0, 50.0, 100.0, 30.0, "Hello", "Alice")
    assert "markup_id" in result
    markups = service.list_markups(sample_pdf)
    assert len(markups) == 1
    assert markups[0]["id"] == result["markup_id"]
    assert markups[0]["author"] == "Alice"


def test_add_callout_returns_markup_id(service, sample_pdf):
    result = service.add_callout(sample_pdf, 1, 50.0, 50.0, "See here", "Bob")
    assert "markup_id" in result
    markups = service.list_markups(sample_pdf)
    assert len(markups) == 1
    assert markups[0]["id"] == result["markup_id"]


def test_add_stamp_returns_markup_id(service, sample_pdf):
    result = service.add_stamp(sample_pdf, 1, "APPROVED", 50.0, 50.0)
    assert "markup_id" in result
    markups = service.list_markups(sample_pdf)
    assert len(markups) == 1
    assert markups[0]["id"] == result["markup_id"]


def test_delete_markup(service, sample_pdf):
    add = service.add_text_box(sample_pdf, 1, 10.0, 10.0, 100.0, 30.0, "To delete", None)
    markup_id = add["markup_id"]

    result = service.delete_markup(sample_pdf, markup_id)
    assert result == {"success": True}
    assert service.list_markups(sample_pdf) == []


def test_delete_markup_raises_if_not_found(service, sample_pdf):
    with pytest.raises(BluebeamDocumentError, match="Markup not found"):
        service.delete_markup(sample_pdf, "99999")


# --- Layer method tests ---

def test_list_layers_empty(service, sample_pdf):
    assert service.list_layers(sample_pdf) == []


def test_add_layer(service, sample_pdf):
    result = service.add_layer(sample_pdf, "Electrical")
    assert result == {"success": True}
    layers = service.list_layers(sample_pdf)
    assert any(l["name"] == "Electrical" for l in layers)


def test_set_layer_visibility(service, sample_pdf):
    service.add_layer(sample_pdf, "Plumbing")
    result = service.set_layer_visibility(sample_pdf, "Plumbing", False)
    assert result == {"success": True}
    layers = service.list_layers(sample_pdf)
    layer = next(l for l in layers if l["name"] == "Plumbing")
    assert layer["visible"] is False


def test_set_layer_visibility_raises_if_not_found(service, sample_pdf):
    with pytest.raises(BluebeamDocumentError, match="Layer not found: Ghost"):
        service.set_layer_visibility(sample_pdf, "Ghost", True)


# --- Workflow method tests ---

def test_flatten_document_removes_annotations(service, sample_pdf):
    service.add_text_box(sample_pdf, 1, 10.0, 10.0, 100.0, 30.0, "Will be removed", None)
    assert len(service.list_markups(sample_pdf)) == 1

    result = service.flatten_document(sample_pdf)
    assert result == {"success": True}
    assert service.list_markups(sample_pdf) == []


def test_export_markup_summary(service, sample_pdf, tmp_path):
    service.add_text_box(sample_pdf, 1, 10.0, 10.0, 100.0, 30.0, "Note", "Alice")
    output = str(tmp_path / "summary.csv")

    result = service.export_markup_summary(sample_pdf, output)

    assert result == {"rows_written": 1}
    assert os.path.exists(output)
    with open(output) as f:
        lines = f.readlines()
    assert len(lines) == 2  # header + 1 row


def test_export_markup_summary_raises_if_output_dir_missing(service, sample_pdf):
    with pytest.raises(BluebeamDocumentError, match="Output directory does not exist"):
        service.export_markup_summary(sample_pdf, "C:\\nonexistent_dir\\out.csv")
