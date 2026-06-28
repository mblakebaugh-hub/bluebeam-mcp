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
