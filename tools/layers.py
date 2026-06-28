from exceptions import BluebeamDocumentError, BluebeamNotAvailableError


def register_layer_tools(mcp, service):
    @mcp.tool()
    def bb_list_layers(path: str) -> list:
        """List all layers and their visibility. Returns [{name, visible}]."""
        try:
            return service.list_layers(path)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_set_layer_visibility(path: str, layer_name: str, visible: bool) -> dict:
        """Show or hide a layer by name."""
        try:
            return service.set_layer_visibility(path, layer_name, visible)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))

    @mcp.tool()
    def bb_add_layer(path: str, layer_name: str) -> dict:
        """Create a new layer in a document."""
        try:
            return service.add_layer(path, layer_name)
        except (BluebeamNotAvailableError, BluebeamDocumentError) as e:
            raise Exception(str(e))
