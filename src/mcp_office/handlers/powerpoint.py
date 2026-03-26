from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

from mcp_office.handlers.base import DocumentHandler
from mcp.types import Tool, TextContent


class PptHandler(DocumentHandler):
    """Handler for PowerPoint operations"""

    def get_tools(self) -> list[Tool]:
        return [
            Tool(
                name="ppt_create",
                description="Create a new PowerPoint presentation",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Save path"},
                        "title": {"type": "string", "description": "Title"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="ppt_add_slide",
                description="Add a slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                        "layout": {"type": "string", "description": "Layout type"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="ppt_add_title",
                description="Add title to slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                        "slide_index": {
                            "type": "integer",
                            "description": "Slide index",
                        },
                        "title": {"type": "string", "description": "Title text"},
                        "font_size": {"type": "integer", "description": "Font size"},
                    },
                    "required": ["path", "slide_index", "title"],
                },
            ),
            Tool(
                name="ppt_add_text",
                description="Add text to slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                        "slide_index": {
                            "type": "integer",
                            "description": "Slide index",
                        },
                        "text": {"type": "string", "description": "Text"},
                        "left": {"type": "number", "description": "Left position"},
                        "top": {"type": "number", "description": "Top position"},
                        "width": {"type": "number", "description": "Width"},
                        "height": {"type": "number", "description": "Height"},
                    },
                    "required": ["path", "slide_index", "text"],
                },
            ),
            Tool(
                name="ppt_add_image",
                description="Add image to slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                        "slide_index": {
                            "type": "integer",
                            "description": "Slide index",
                        },
                        "image_path": {"type": "string", "description": "Image path"},
                        "left": {"type": "number", "description": "Left"},
                        "top": {"type": "number", "description": "Top"},
                        "width": {"type": "number", "description": "Width"},
                    },
                    "required": ["path", "slide_index", "image_path"],
                },
            ),
            Tool(
                name="ppt_add_table",
                description="Add table to slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                        "slide_index": {
                            "type": "integer",
                            "description": "Slide index",
                        },
                        "data": {
                            "type": "array",
                            "items": {"type": "array", "items": {"type": "string"}},
                        },
                        "left": {"type": "number", "description": "Left"},
                        "top": {"type": "number", "description": "Top"},
                    },
                    "required": ["path", "slide_index", "data"],
                },
            ),
            Tool(
                name="ppt_list_slides",
                description="List slides",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Presentation path"},
                    },
                    "required": ["path"],
                },
            ),
        ]

    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        try:
            if tool_name == "ppt_create":
                return self._create(arguments)
            elif tool_name == "ppt_add_slide":
                return self._add_slide(arguments)
            elif tool_name == "ppt_add_title":
                return self._add_title(arguments)
            elif tool_name == "ppt_add_text":
                return self._add_text(arguments)
            elif tool_name == "ppt_add_image":
                return self._add_image(arguments)
            elif tool_name == "ppt_add_table":
                return self._add_table(arguments)
            elif tool_name == "ppt_list_slides":
                return self._list_slides(arguments)
            return self.error_result(f"Unknown tool: {tool_name}")
        except Exception as e:
            return self.error_result(str(e))

    def _create(self, args: dict) -> list[TextContent]:
        prs = Presentation()
        if title := args.get("title"):
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = title
        prs.save(args["path"])
        return self.success_result(f"Presentation created: {args['path']}")

    def _add_slide(self, args: dict) -> list[TextContent]:
        prs = Presentation(args["path"])
        layout_map = {"title": 0, "title_content": 1, "blank": 6}
        idx = layout_map.get(args.get("layout", "blank"), 6)
        prs.slides.add_slide(prs.slide_layouts[idx])
        prs.save(args["path"])
        return self.success_result(f"Added slide at index {len(prs.slides) - 1}")

    def _add_title(self, args: dict) -> list[TextContent]:
        prs = Presentation(args["path"])
        slide = prs.slides[args["slide_index"]]
        title = slide.shapes.title
        title.text = args["title"]
        if fs := args.get("font_size"):
            for p in title.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = PptxPt(fs)
        prs.save(args["path"])
        return self.success_result("Added title to slide")

    def _add_text(self, args: dict) -> list[TextContent]:
        prs = Presentation(args["path"])
        slide = prs.slides[args["slide_index"]]
        left = PptxInches(args.get("left", 1))
        top = PptxInches(args.get("top", 1))
        width = PptxInches(args.get("width", 6))
        height = PptxInches(args.get("height", 1))
        txBox = slide.shapes.add_textbox(left, top, width, height)
        txBox.text_frame.text = args["text"]
        prs.save(args["path"])
        return self.success_result("Added text to slide")

    def _add_image(self, args: dict) -> list[TextContent]:
        prs = Presentation(args["path"])
        slide = prs.slides[args["slide_index"]]
        left = PptxInches(args.get("left", 1))
        top = PptxInches(args.get("top", 1))
        width = PptxInches(args.get("width", 4))
        slide.shapes.add_picture(args["image_path"], left, top, width=width)
        prs.save(args["path"])
        return self.success_result("Added image to slide")

    def _add_table(self, args: dict) -> list[TextContent]:

        prs = Presentation(args["path"])
        slide = prs.slides[args["slide_index"]]
        data = args["data"]
        rows, cols = len(data), len(data[0])
        left = PptxInches(args.get("left", 1))
        top = PptxInches(args.get("top", 2))
        table = slide.shapes.add_table(rows, cols, left, top).table
        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                table.cell(r_idx, c_idx).text = val
        prs.save(args["path"])
        return self.success_result(f"Added table {rows}x{cols}")

    def _list_slides(self, args: dict) -> list[TextContent]:
        import json

        prs = Presentation(args["path"])
        slides = [{"index": i} for i in range(len(prs.slides))]
        return [TextContent(type="text", text=json.dumps(slides, indent=2))]
