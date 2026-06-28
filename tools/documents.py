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
        """Acknowledge close request. Note: cannot force-close a file in Revu via API."""
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
        """List open PDFs. Note: returns empty — Revu.Launcher does not expose a document list."""
        try:
            return service.list_open_documents()
        except BluebeamNotAvailableError as e:
            raise Exception(str(e))
