from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from mcp_office.handlers import WordHandler, ExcelHandler, PptHandler, PdfHandler


server = Server("mcp-office")

_handlers = [
    WordHandler(),
    ExcelHandler(),
    PptHandler(),
    PdfHandler(),
]


def _get_tool_map():
    tool_map = {}
    for handler in _handlers:
        for tool in handler.get_tools():
            tool_map[tool.name] = handler
    return tool_map


_tool_map = _get_tool_map()


@server.list_tools()
async def list_tools() -> list[Tool]:
    tools = []
    for handler in _handlers:
        tools.extend(handler.get_tools())
    return tools


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    handler = _tool_map.get(name)
    if handler:
        return await handler.execute(name, arguments)
    return [TextContent(type="text", text=f"Unknown tool: {name}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )
