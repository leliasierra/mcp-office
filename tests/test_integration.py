import pytest
import json
from pathlib import Path
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from mcp_office.handlers import WordHandler, ExcelHandler, PptHandler, PdfHandler
from mcp_office.handlers.base import DocumentHandler


class TestIntegration:
    """Integration tests - handlers working together"""

    def test_all_handlers_implement_interface(self):
        """Verify all handlers implement DocumentHandler"""
        handlers = [
            WordHandler(),
            ExcelHandler(),
            PptHandler(),
            PdfHandler(),
        ]
        for handler in handlers:
            assert isinstance(handler, DocumentHandler)
            assert hasattr(handler, "get_tools")
            assert hasattr(handler, "execute")

    def test_word_handler_tools_unique(self):
        """Word handler has unique tool names"""
        word = WordHandler()
        tools = word.get_tools()
        tool_names = [t.name for t in tools]
        assert len(tool_names) == len(set(tool_names))

    def test_excel_handler_tools_unique(self):
        """Excel handler has unique tool names"""
        excel = ExcelHandler()
        tools = excel.get_tools()
        tool_names = [t.name for t in tools]
        assert len(tool_names) == len(set(tool_names))

    def test_ppt_handler_tools_unique(self):
        """PPT handler has unique tool names"""
        ppt = PptHandler()
        tools = ppt.get_tools()
        tool_names = [t.name for t in tools]
        assert len(tool_names) == len(set(tool_names))

    def test_pdf_handler_tools_unique(self):
        """PDF handler has unique tool names"""
        pdf = PdfHandler()
        tools = pdf.get_tools()
        tool_names = [t.name for t in tools]
        assert len(tool_names) == len(set(tool_names))

    def test_no_tool_name_conflicts_across_handlers(self):
        """No tool name conflicts between handlers"""
        word = WordHandler()
        excel = ExcelHandler()
        ppt = PptHandler()
        pdf = PdfHandler()

        all_tools = (
            word.get_tools() + excel.get_tools() + ppt.get_tools() + pdf.get_tools()
        )
        tool_names = [t.name for t in all_tools]
        assert len(tool_names) == len(set(tool_names)), "Tool name conflict detected"

    @pytest.mark.asyncio
    async def test_word_to_pdf_conversion_flow(self, tmp_path):
        """Test Word -> PDF conversion (requires LibreOffice or skip)"""
        pytest.skip("Requires LibreOffice or Microsoft Word")

    @pytest.mark.asyncio
    async def test_pdf_to_word_conversion_flow(self, tmp_path):
        """Test PDF -> Word conversion (requires LibreOffice or skip)"""
        pytest.skip("Requires LibreOffice or Microsoft Word")
        """Test Excel read/write cycle"""
        excel = ExcelHandler()
        xlsx_path = tmp_path / "data.xlsx"

        await excel.execute("excel_create", {"path": str(xlsx_path)})
        await excel.execute(
            "excel_write_cell",
            {"path": str(xlsx_path), "sheet": "Sheet", "cell": "A1", "value": "Test"},
        )
        await excel.execute(
            "excel_add_row",
            {"path": str(xlsx_path), "sheet": "Sheet", "data": ["Row1", "Row2"]},
        )

        result = await excel.execute("excel_read", {"path": str(xlsx_path)})
        data = json.loads(result[0].text)
        assert len(data) > 0

    @pytest.mark.asyncio
    async def test_ppt_generate(self, tmp_path):
        """Test PPT generation from SVGs"""
        ppt = PptHandler()
        svg_dir = tmp_path / "slides"
        svg_dir.mkdir()
        svg = svg_dir / "slide.svg"
        svg.write_text(
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">'
            '<rect width="1280" height="720" fill="#fff"/>'
            '<text x="640" y="360" text-anchor="middle" font-size="36">Hello</text></svg>'
        )

        output = tmp_path / "out.pptx"
        result = await ppt.execute(
            "ppt_generate", {"svg_source": str(svg_dir), "output": str(output)}
        )
        assert "generated" in result[0].text.lower()
        assert output.exists()

    def test_handler_error_results(self):
        """All handlers return proper error format"""
        handlers = [WordHandler(), ExcelHandler(), PptHandler(), PdfHandler()]

        for handler in handlers:
            error_result = handler.error_result("test error")
            assert error_result[0].text.startswith("Error:")
            assert "test error" in error_result[0].text

    def test_handler_success_results(self):
        """All handlers return proper success format"""
        handlers = [WordHandler(), ExcelHandler(), PptHandler(), PdfHandler()]

        for handler in handlers:
            success_result = handler.success_result("test success")
            assert "test success" in success_result[0].text
