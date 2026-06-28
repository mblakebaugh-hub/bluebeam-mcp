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
