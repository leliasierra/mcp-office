from mcp_office.handlers.base import DocumentHandler
from mcp_office.handlers.word import WordHandler
from mcp_office.handlers.excel import ExcelHandler
from mcp_office.handlers.powerpoint import PptHandler
from mcp_office.handlers.pdf import PdfHandler
from mcp_office.handlers.ocr import OcrHandler

__all__ = [
    "DocumentHandler",
    "WordHandler",
    "ExcelHandler",
    "PptHandler",
    "PdfHandler",
    "OcrHandler",
]
