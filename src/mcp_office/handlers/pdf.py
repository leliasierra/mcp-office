import fitz
from pathlib import Path
from pypdf import PdfReader
from docx import Document as DocxDocument
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
import img2pdf

from mcp_office.handlers.base import DocumentHandler
from mcp_office.spec import SpecManager
from mcp.types import Tool, TextContent


class PdfHandler(DocumentHandler):
    """Handler for PDF operations"""

    def get_tools(self) -> list[Tool]:
        return [
            Tool(
                name="pdf_read",
                description="Read PDF content",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "extract_text": {
                            "type": "boolean",
                            "description": "Extract text",
                        },
                        "page_start": {"type": "integer", "description": "Start page"},
                        "page_end": {"type": "integer", "description": "End page"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="pdf_get_info",
                description="Get PDF metadata",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="pdf_extract_images",
                description="Extract images from PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "output_dir": {
                            "type": "string",
                            "description": "Output directory",
                        },
                    },
                    "required": ["path", "output_dir"],
                },
            ),
            Tool(
                name="pdf_split",
                description="Split PDF into pages",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "output_dir": {
                            "type": "string",
                            "description": "Output directory",
                        },
                    },
                    "required": ["path", "output_dir"],
                },
            ),
            Tool(
                name="pdf_merge",
                description="Merge multiple PDFs",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "sources": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "PDF paths",
                        },
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["sources", "output"],
                },
            ),
            Tool(
                name="pdf_rotate",
                description="Rotate PDF pages",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "degrees": {"type": "integer", "description": "Degrees"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path", "degrees"],
                },
            ),
            Tool(
                name="pdf_compress",
                description="Compress PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="pdf_add_watermark",
                description="Add watermark to PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "text": {"type": "string", "description": "Watermark text"},
                        "opacity": {"type": "number", "description": "Opacity"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path", "text"],
                },
            ),
            Tool(
                name="pdf_add_page_numbers",
                description="Add page numbers",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="pdf_protect",
                description="Protect PDF with password",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "password": {"type": "string", "description": "Password"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path", "password"],
                },
            ),
            Tool(
                name="pdf_unprotect",
                description="Remove PDF protection",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "password": {"type": "string", "description": "Password"},
                        "output": {"type": "string", "description": "Output path"},
                    },
                    "required": ["path", "password"],
                },
            ),
            Tool(
                name="pdf_extract_text",
                description="Extract text from pages",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "page_start": {"type": "integer", "description": "Start page"},
                        "page_end": {"type": "integer", "description": "End page"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="word_to_pdf",
                description="Convert Word to PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "input_path": {"type": "string", "description": "Word path"},
                        "output_path": {"type": "string", "description": "PDF path"},
                    },
                    "required": ["input_path", "output_path"],
                },
            ),
            Tool(
                name="pdf_to_word",
                description="Convert PDF to Word",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "input_path": {"type": "string", "description": "PDF path"},
                        "output_path": {"type": "string", "description": "Word path"},
                    },
                    "required": ["input_path", "output_path"],
                },
            ),
            Tool(
                name="pdf_create",
                description="Create PDF with ReportLab",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Output path"},
                        "title": {"type": "string", "description": "Document title"},
                        "page_size": {
                            "type": "string",
                            "description": "Page size: letter or A4",
                        },
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="pdf_add_text",
                description="Add text to PDF (ReportLab)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "PDF path"},
                        "text": {"type": "string", "description": "Text to add"},
                        "x": {"type": "number", "description": "X position"},
                        "y": {"type": "number", "description": "Y position"},
                        "font_size": {"type": "number", "description": "Font size"},
                    },
                    "required": ["path", "text"],
                },
            ),
            Tool(
                name="pdf_images_to_pdf",
                description="Convert images to PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "image_paths": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Image file paths",
                        },
                        "output": {"type": "string", "description": "Output PDF path"},
                    },
                    "required": ["image_paths", "output"],
                },
            ),
        ]

    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        try:
            if tool_name == "pdf_read":
                return self._read(arguments)
            elif tool_name == "pdf_get_info":
                return self._get_info(arguments)
            elif tool_name == "pdf_extract_images":
                return self._extract_images(arguments)
            elif tool_name == "pdf_split":
                return self._split(arguments)
            elif tool_name == "pdf_merge":
                return self._merge(arguments)
            elif tool_name == "pdf_rotate":
                return self._rotate(arguments)
            elif tool_name == "pdf_compress":
                return self._compress(arguments)
            elif tool_name == "pdf_add_watermark":
                return self._add_watermark(arguments)
            elif tool_name == "pdf_add_page_numbers":
                return self._add_page_numbers(arguments)
            elif tool_name == "pdf_protect":
                return self._protect(arguments)
            elif tool_name == "pdf_unprotect":
                return self._unprotect(arguments)
            elif tool_name == "pdf_extract_text":
                return self._extract_text(arguments)
            elif tool_name == "word_to_pdf":
                return self._word_to_pdf(arguments)
            elif tool_name == "pdf_to_word":
                return self._pdf_to_word(arguments)
            elif tool_name == "pdf_create":
                return self._pdf_create(arguments)
            elif tool_name == "pdf_add_text":
                return self._pdf_add_text(arguments)
            elif tool_name == "pdf_images_to_pdf":
                return self._images_to_pdf(arguments)
            return self.error_result(f"Unknown tool: {tool_name}")
        except Exception as e:
            return self.error_result(str(e))

    def _read(self, args: dict) -> list[TextContent]:
        import json

        doc = fitz.open(args["path"])
        result = {"pages": len(doc), "metadata": doc.metadata}
        if args.get("extract_text"):
            start = args.get("page_start", 1) - 1
            end = args.get("page_end", len(doc))
            text = "\n".join(
                doc[i].get_text() for i in range(start, min(end, len(doc)))
            )
            result["text"] = text[:10000]
        doc.close()
        return [TextContent(type="text", text=json.dumps(result, indent=2))]

    def _get_info(self, args: dict) -> list[TextContent]:
        import json

        doc = fitz.open(args["path"])
        info = {
            "pages": len(doc),
            "title": doc.metadata.get("title", ""),
            "author": doc.metadata.get("author", ""),
            "creator": doc.metadata.get("creator", ""),
        }
        doc.close()
        return [TextContent(type="text", text=json.dumps(info, indent=2))]

    def _extract_images(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        output_dir = Path(args["output_dir"])
        output_dir.mkdir(parents=True, exist_ok=True)
        count = 0
        for page_num, page in enumerate(doc):
            for img_num, img in enumerate(page.get_images()):
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pix.save(output_dir / f"page{page_num + 1}_img{img_num + 1}.png")
                count += 1
        doc.close()
        return self.success_result(f"Extracted {count} images")

    def _split(self, args: dict) -> list[TextContent]:
        from pathlib import Path

        doc = fitz.open(args["path"])
        output_dir = Path(args["output_dir"])
        output_dir.mkdir(parents=True, exist_ok=True)
        for i in range(len(doc)):
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            new_doc.save(output_dir / f"page{i + 1}.pdf")
        doc.close()
        return self.success_result(f"Split into {len(doc)} pages")

    def _merge(self, args: dict) -> list[TextContent]:
        output = fitz.open()
        for src in args["sources"]:
            output.insert_pdf(fitz.open(src))
        output.save(args["output"])
        output.close()
        return self.success_result(f"Merged {len(args['sources'])} PDFs")

    def _rotate(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        for page in doc:
            page.set_rotation(page.rotation + args["degrees"])
        output = args.get("output", args["path"])
        doc.save(output)
        doc.close()
        return self.success_result(f"Rotated {args['degrees']} degrees")

    def _compress(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        output = args.get("output", args["path"])
        doc.save(output, garbage=4, deflate=True, clean=True)
        doc.close()
        return self.success_result(f"Compressed to {output}")

    def _add_watermark(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        text = args["text"]
        opacity = args.get("opacity", 0.3)
        for page in doc:
            page.insert_text((100, 100), text, opacity=opacity, fontsize=72)
        output = args.get("output", args["path"])
        doc.save(output)
        doc.close()
        return self.success_result(f"Added watermark: {text}")

    def _add_page_numbers(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        for i, page in enumerate(doc):
            page.insert_text((500, 800), str(i + 1), fontsize=12)
        output = args.get("output", args["path"])
        doc.save(output)
        doc.close()
        return self.success_result("Added page numbers")

    def _protect(self, args: dict) -> list[TextContent]:
        import pikepdf
        from pathlib import Path

        input_path = Path(args["path"])
        output = args.get("output")

        if not output:
            output = str(
                input_path.parent / f"{input_path.stem}_protected{input_path.suffix}"
            )

        pdf = pikepdf.open(args["path"])
        pdf.save(
            output,
            encryption=pikepdf.Encryption(
                owner=args["password"], user=args["password"]
            ),
        )
        return self.success_result(f"PDF protected with password: {output}")

    def _unprotect(self, args: dict) -> list[TextContent]:
        import pikepdf

        try:
            pdf = pikepdf.open(args["path"], password=args["password"])
            output = args.get("output", args["path"].replace(".pdf", "_unlocked.pdf"))
            pdf.save(output)
            return self.success_result(f"PDF unlocked: {output}")
        except pikepdf.PasswordError:
            return self.error_result("Invalid password")
        except Exception as e:
            return self.error_result(f"Failed to unlock PDF: {str(e)}")

    def _extract_text(self, args: dict) -> list[TextContent]:
        doc = fitz.open(args["path"])
        start = args.get("page_start", 1) - 1
        end = args.get("page_end", len(doc))
        text = "\n".join(doc[i].get_text() for i in range(start, min(end, len(doc))))
        doc.close()
        return [TextContent(type="text", text=text)]

    def _word_to_pdf(self, args: dict) -> list[TextContent]:
        from docx2pdf import convert

        convert(args["input_path"], args["output_path"])
        SpecManager("word", args["input_path"]).add_post_processing(
            "convert_to_pdf", output=args["output_path"]
        )
        return self.success_result(f"Converted to {args['output_path']}")

    def _pdf_to_word(self, args: dict) -> list[TextContent]:
        pdf = PdfReader(args["input_path"])
        doc = DocxDocument()
        for page in pdf.pages:
            if text := page.extract_text():
                doc.add_paragraph(text)
                doc.add_page_break()
        doc.save(args["output_path"])
        SpecManager("pdf", args["input_path"]).add_post_processing(
            "convert_to_word", output=args["output_path"]
        )
        return self.success_result(f"Converted to {args['output_path']}")

    def _pdf_create(self, args: dict) -> list[TextContent]:
        page_size = A4 if args.get("page_size", "").upper() == "A4" else letter
        c = canvas.Canvas(args["path"], pagesize=page_size)
        if title := args.get("title"):
            c.setTitle(title)
        c.save()
        return self.success_result(f"PDF created: {args['path']}")

    def _pdf_add_text(self, args: dict) -> list[TextContent]:
        c = canvas.Canvas(args["path"], pagesize=letter)
        x = args.get("x", 100)
        y = args.get("y", 700)
        font_size = args.get("font_size", 12)
        c.setFont("Helvetica", font_size)
        c.drawString(x, y, args["text"])
        c.save()
        return self.success_result("Text added to PDF")

    def _images_to_pdf(self, args: dict) -> list[TextContent]:
        with open(args["output"], "wb") as f:
            f.write(img2pdf.convert(args["image_paths"]))
        return self.success_result(f"Images converted to PDF: {args['output']}")
