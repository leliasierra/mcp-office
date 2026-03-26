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
    async def test_ppt_slide_management(self, tmp_path):
        """Test PPT slide creation and listing"""
        ppt = PptHandler()
        pptx_path = tmp_path / "presentation.pptx"

        await ppt.execute("ppt_create", {"path": str(pptx_path), "title": "Test"})
        await ppt.execute(
            "ppt_add_slide", {"path": str(pptx_path), "layout": "title_content"}
        )
        await ppt.execute("ppt_add_slide", {"path": str(pptx_path), "layout": "blank"})

        result = await ppt.execute("ppt_list_slides", {"path": str(pptx_path)})
        slides = json.loads(result[0].text)
        assert len(slides) == 3

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
