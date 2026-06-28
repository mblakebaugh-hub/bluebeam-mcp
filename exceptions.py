class BluebeamNotAvailableError(Exception):
    """Revu is not installed, not running, or connection was lost."""


class BluebeamDocumentError(Exception):
    """File not found, page out of range, markup/layer not found."""
