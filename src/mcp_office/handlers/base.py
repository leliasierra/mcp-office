from abc import ABC, abstractmethod
from mcp.types import Tool, TextContent


class DocumentHandler(ABC):
    """Abstract base class for document handlers - ISP + OCP"""

    @abstractmethod
    def get_tools(self) -> list[Tool]:
        """Return list of tools for this handler"""
        pass

    @abstractmethod
    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        """Execute a tool and return results"""
        pass

    @staticmethod
    def error_result(message: str) -> list[TextContent]:
        return [TextContent(type="text", text=f"Error: {message}")]

    @staticmethod
    def success_result(message: str) -> list[TextContent]:
        return [TextContent(type="text", text=message)]
