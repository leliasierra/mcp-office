import pytesseract
from pathlib import Path
from mcp_office.handlers.base import DocumentHandler
from mcp.types import Tool, TextContent

# Configure Tesseract path for Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class OcrHandler(DocumentHandler):
    """Handler for OCR operations using Tesseract"""

    def get_tools(self) -> list[Tool]:
        return [
            Tool(
                name="ocr_image",
                description="Extract text from image using Tesseract OCR",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Image file path"},
                        "language": {
                            "type": "string",
                            "description": "Language code (eng, spa, etc.)",
                            "default": "eng",
                        },
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="ocr_image_to_text_file",
                description="Save OCR text from image to file",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Image file path"},
                        "output": {"type": "string", "description": "Output text file path"},
                        "language": {
                            "type": "string",
                            "description": "Language code (eng, spa, etc.)",
                            "default": "eng",
                        },
                    },
                    "required": ["path", "output"],
                },
            ),
            Tool(
                name="ocr_pdf",
                description="Extract text from PDF using Tesseract (for scanned PDFs)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF file path"},
                        "language": {
                            "type": "string",
                            "description": "Language code (eng, spa, etc.)",
                            "default": "eng",
                        },
                        "dpi": {
                            "type": "integer",
                            "description": "DPI for image conversion",
                            "default": 300,
                        },
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="ocr_list_languages",
                description="List available Tesseract OCR languages",
                inputSchema={
                    "type": "object",
                    "properties": {},
                },
            ),
        ]

    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        try:
            if tool_name == "ocr_image":
                return self._ocr_image(arguments)
            elif tool_name == "ocr_image_to_text_file":
                return self._ocr_image_to_file(arguments)
            elif tool_name == "ocr_pdf":
                return self._ocr_pdf(arguments)
            elif tool_name == "ocr_list_languages":
                return self._list_languages(arguments)
            return self.error_result(f"Unknown tool: {tool_name}")
        except Exception as e:
            return self.error_result(str(e))

    def _ocr_image(self, args: dict) -> list[TextContent]:
        from PIL import Image

        img_path = args["path"]
        lang = args.get("language", "eng")

        img = Image.open(img_path)
        text = pytesseract.image_to_string(img, lang=lang)
        return [TextContent(type="text", text=text)]

    def _ocr_image_to_file(self, args: dict) -> list[TextContent]:
        from PIL import Image

        img_path = args["path"]
        output_path = args["output"]
        lang = args.get("language", "eng")

        img = Image.open(img_path)
        text = pytesseract.image_to_string(img, lang=lang)

        Path(output_path).write_text(text, encoding="utf-8")
        return self.success_result(f"OCR text saved to {output_path}")

    def _ocr_pdf(self, args: dict) -> list[TextContent]:
        import fitz
        from PIL import Image
        import io

        pdf_path = args["path"]
        lang = args.get("language", "eng")
        dpi = args.get("dpi", 300)

        doc = fitz.open(pdf_path)
        all_text = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img, lang=lang)
            all_text.append(f"--- Page {page_num + 1} ---\n{text}")

        doc.close()
        return [TextContent(type="text", text="\n\n".join(all_text))]

    def _list_languages(self, args: dict) -> list[TextContent]:
        langs = pytesseract.get_languages()
        return [TextContent(type="text", text="\n".join(sorted(langs)))]
