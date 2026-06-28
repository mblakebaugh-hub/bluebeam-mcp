import pytest
import fitz
from unittest.mock import MagicMock


class SyncDispatcher:
    """Runs COM-thread callables synchronously for unit testing."""

    def run(self, fn, timeout=30):
        return fn()


@pytest.fixture
def mock_launcher():
    return MagicMock()


@pytest.fixture
def service(mock_launcher):
    from bluebeam_service import BluebeamService
    return BluebeamService(_launcher=mock_launcher, _com=SyncDispatcher())


@pytest.fixture
def sample_pdf(tmp_path):
    """Create a minimal single-page PDF for testing."""
    path = str(tmp_path / "sample.pdf")
    doc = fitz.open()
    doc.new_page(width=612, height=792)
    doc.save(path)
    doc.close()
    return path
