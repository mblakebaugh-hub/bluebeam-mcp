from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_workflow_tools(mcp, service):
    @mcp.tool()
    def bb_flatten_document(path: str) -> dict:
        """Flatten all markups into the PDF content (irreversible)."""
        try:
            return service.flatten_document(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_export_markup_summary(path: str, output_path: str) -> dict:
        """Export a CSV markup summary report. Returns rows_written count."""
        try:
            return service.export_markup_summary(path, output_path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
