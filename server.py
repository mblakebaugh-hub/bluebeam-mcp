from mcp.server.fastmcp import FastMCP

from bluebeam_service import BluebeamService
from tools.documents import register_document_tools
from tools.layers import register_layer_tools
from tools.markups import register_markup_tools
from tools.workflows import register_workflow_tools

mcp = FastMCP("bluebeam-mcp")
service = BluebeamService()

register_document_tools(mcp, service)
register_markup_tools(mcp, service)
register_layer_tools(mcp, service)
register_workflow_tools(mcp, service)

if __name__ == "__main__":
    mcp.run()
