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
