import pytest
import asyncio
import json
import sys
from pathlib import Path


class TestE2E:
    """End-to-End tests - Full MCP server flow"""

    @pytest.mark.asyncio
    async def test_mcp_server_lists_all_tools(self):
        """Test MCP server returns all tools from all handlers"""
        from mcp_office import list_tools

        tools = await list_tools()
        tool_names = [t.name for t in tools]

        assert "create_document" in tool_names
        assert "excel_create" in tool_names
        assert "ppt_create" in tool_names
        assert "pdf_read" in tool_names
        assert len(tool_names) > 30

    @pytest.mark.asyncio
    async def test_mcp_execute_word_flow(self, tmp_path):
        """Test complete Word document creation flow"""
        from mcp_office import call_tool

        doc_path = tmp_path / "report.docx"

        await call_tool(
            "create_document", {"path": str(doc_path), "title": "Q1 Report"}
        )
        await call_tool(
            "add_heading", {"path": str(doc_path), "text": "Introduction", "level": 1}
        )
        await call_tool(
            "add_paragraph",
            {"path": str(doc_path), "text": "This is the quarterly report."},
        )
        await call_tool(
            "add_table",
            {
                "path": str(doc_path),
                "data": [["Metric", "Value"], ["Sales", "1000"]],
                "header_row": True,
            },
        )

        result = await call_tool(
            "open_document", {"path": str(doc_path), "extract_text": True}
        )

        data = json.loads(result[0].text)
        assert data["paragraphs"] >= 2
        assert "quarterly report" in data["text"].lower()

    @pytest.mark.asyncio
    async def test_mcp_execute_excel_flow(self, tmp_path):
        """Test complete Excel workflow"""
        from mcp_office import call_tool

        xlsx_path = tmp_path / "sales.xlsx"

        await call_tool("excel_create", {"path": str(xlsx_path), "sheet_name": "Sales"})
        await call_tool(
            "excel_write_cell",
            {
                "path": str(xlsx_path),
                "sheet": "Sales",
                "cell": "A1",
                "value": "Product",
            },
        )
        await call_tool(
            "excel_write_cell",
            {"path": str(xlsx_path), "sheet": "Sales", "cell": "B1", "value": "Amount"},
        )
        await call_tool(
            "excel_add_row",
            {"path": str(xlsx_path), "sheet": "Sales", "data": ["Widget A", "500"]},
        )

        result = await call_tool("excel_read", {"path": str(xlsx_path)})
        data = json.loads(result[0].text)
        assert len(data) >= 2

    @pytest.mark.asyncio
    async def test_mcp_execute_ppt_flow(self, tmp_path):
        """Test complete PowerPoint workflow"""
        from mcp_office import call_tool

        pptx_path = tmp_path / "presentation.pptx"

        await call_tool(
            "ppt_create", {"path": str(pptx_path), "title": "Company Overview"}
        )
        await call_tool(
            "ppt_add_slide", {"path": str(pptx_path), "layout": "title_content"}
        )
        await call_tool(
            "ppt_add_title",
            {"path": str(pptx_path), "slide_index": 0, "title": "Welcome"},
        )

        result = await call_tool("ppt_list_slides", {"path": str(pptx_path)})
        slides = json.loads(result[0].text)
        assert len(slides) >= 1

    @pytest.mark.asyncio
    async def test_mcp_unknown_tool_returns_error(self):
        """Test unknown tool returns proper error"""
        from mcp_office import call_tool

        result = await call_tool("nonexistent_tool", {})
        assert "unknown tool" in result[0].text.lower()

    @pytest.mark.asyncio
    async def test_mcp_invalid_path_returns_error(self):
        """Test invalid file path returns error"""
        from mcp_office import call_tool

        result = await call_tool("open_document", {"path": "/nonexistent/path.docx"})
        assert "error" in result[0].text.lower()

    @pytest.mark.asyncio
    async def test_word_find_replace_flow(self, tmp_path):
        """Test find and replace workflow"""
        from mcp_office import call_tool

        doc_path = tmp_path / "document.docx"

        await call_tool("create_document", {"path": str(doc_path)})
        await call_tool("add_paragraph", {"path": str(doc_path), "text": "Hello World"})

        result = await call_tool(
            "find_replace",
            {"path": str(doc_path), "find": "World", "replace": "Python"},
        )

        assert "replaced" in result[0].text.lower()

    @pytest.mark.asyncio
    async def test_excel_formula_flow(self, tmp_path):
        """Test Excel formula workflow"""
        from mcp_office import call_tool

        xlsx_path = tmp_path / "formulas.xlsx"

        await call_tool("excel_create", {"path": str(xlsx_path)})
        await call_tool(
            "excel_write_cell",
            {"path": str(xlsx_path), "sheet": "Sheet", "cell": "A1", "value": "10"},
        )
        await call_tool(
            "excel_write_cell",
            {"path": str(xlsx_path), "sheet": "Sheet", "cell": "A2", "value": "20"},
        )
        await call_tool(
            "excel_add_formula",
            {
                "path": str(xlsx_path),
                "sheet": "Sheet",
                "cell": "A3",
                "formula": "=SUM(A1:A2)",
            },
        )

        result = await call_tool("excel_read", {"path": str(xlsx_path)})
        data = json.loads(result[0].text)
        assert len(data) == 3

    @pytest.mark.asyncio
    async def test_pdf_read_flow(self, tmp_path):
        """Test PDF read operation"""
        pytest.skip("Requires Word installed for conversion")

    @pytest.mark.asyncio
    async def test_all_conversions_available(self):
        """Test all conversion tools are available"""
        from mcp_office import list_tools

        tools = await list_tools()
        tool_names = [t.name for t in tools]

        assert "word_to_pdf" in tool_names
        assert "pdf_to_word" in tool_names
        assert "excel_to_pdf" in tool_names

    @pytest.mark.asyncio
    async def test_handler_dispatch_correctly(self):
        """Test tools are dispatched to correct handler"""
        from mcp_office import call_tool

        result = await call_tool("excel_list_sheets", {"path": "/nonexistent.xlsx"})
        assert "error" in result[0].text.lower() or "created" in result[0].text.lower()


class TestContract:
    """Contract tests - validate tool schemas"""

    @pytest.mark.asyncio
    async def test_all_tools_have_schemas(self):
        """All tools must have inputSchema"""
        from mcp_office import list_tools

        tools = await list_tools()

        for tool in tools:
            assert tool.inputSchema is not None
            assert "type" in tool.inputSchema

    @pytest.mark.asyncio
    async def test_all_tools_have_required_fields(self):
        """All tools with required fields are properly defined"""
        from mcp_office import list_tools

        tools = await list_tools()

        for tool in tools:
            props = tool.inputSchema.get("properties", {})
            required = tool.inputSchema.get("required", [])
            for req_field in required:
                assert req_field in props, f"{tool.name} missing {req_field}"

    @pytest.mark.asyncio
    async def test_tool_names_follow_convention(self):
        """Tool names follow naming convention"""
        from mcp_office import list_tools

        tools = await list_tools()

        for tool in tools:
            assert "_" in tool.name or tool.name.islower()
            assert not any(c.isupper() for c in tool.name)

    @pytest.mark.asyncio
    async def test_descriptions_are_non_empty(self):
        """All tools have descriptions"""
        from mcp_office import list_tools

        tools = await list_tools()

        for tool in tools:
            assert tool.description
            assert len(tool.description) > 5
