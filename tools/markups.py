from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_markup_tools(mcp, service):
    @mcp.tool()
    def bb_list_markups(path: str, page: int = None) -> list:
        """List markups in a document. Omit page to list all pages.
        Returns list of {id, type, page, author, subject, comment, date, x, y}."""
        try:
            return service.list_markups(path, page)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_text_box(path: str, page: int, x: float, y: float,
                        width: float, height: float, text: str,
                        author: str = None) -> dict:
        """Add a text box markup. x, y, width, height in PDF points (1/72 in), origin bottom-left."""
        try:
            return service.add_text_box(path, page, x, y, width, height, text, author)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_callout(path: str, page: int, x: float, y: float,
                       text: str, author: str = None) -> dict:
        """Add a callout (leader line + text box). x, y in PDF points, origin bottom-left."""
        try:
            return service.add_callout(path, page, x, y, text, author)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_stamp(path: str, page: int, stamp_name: str,
                     x: float, y: float) -> dict:
        """Apply a named stamp. stamp_name must match a stamp in Revu's stamp library."""
        try:
            return service.add_stamp(path, page, stamp_name, x, y)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_delete_markup(path: str, markup_id: str) -> dict:
        """Delete a markup by its ID (get IDs from bb_list_markups)."""
        try:
            return service.delete_markup(path, markup_id)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
