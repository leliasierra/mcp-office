from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from mcp_office.base import hex_to_rgb, rgb_to_hex
from mcp_office.handlers.base import DocumentHandler
from mcp.types import Tool, TextContent


class ExcelHandler(DocumentHandler):
    """Handler for Excel operations"""

    def get_tools(self) -> list[Tool]:
        return [
            Tool(
                name="excel_create",
                description="Create a new Excel workbook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {
                            "type": "string",
                            "description": "Path to save workbook",
                        },
                        "sheet_name": {"type": "string", "description": "Sheet name"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="excel_read",
                description="Read data from Excel workbook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="excel_write_cell",
                description="Write value to a cell",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "cell": {"type": "string", "description": "Cell (e.g., A1)"},
                        "value": {"type": "string", "description": "Value"},
                    },
                    "required": ["path", "sheet", "cell", "value"],
                },
            ),
            Tool(
                name="excel_add_row",
                description="Add a row of data",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "data": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Row data",
                        },
                    },
                    "required": ["path", "sheet", "data"],
                },
            ),
            Tool(
                name="excel_add_formula",
                description="Add a formula to a cell",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "cell": {"type": "string", "description": "Cell"},
                        "formula": {"type": "string", "description": "Formula"},
                    },
                    "required": ["path", "sheet", "cell", "formula"],
                },
            ),
            Tool(
                name="excel_format_cell",
                description="Format a cell",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "cell": {"type": "string", "description": "Cell"},
                        "bold": {"type": "boolean", "description": "Bold"},
                        "font_size": {"type": "integer", "description": "Font size"},
                        "font_color": {
                            "type": "string",
                            "description": "Font color hex",
                        },
                        "bg_color": {
                            "type": "string",
                            "description": "Background color",
                        },
                    },
                    "required": ["path", "sheet", "cell"],
                },
            ),
            Tool(
                name="excel_add_table",
                description="Add a table to a sheet",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "data": {
                            "type": "array",
                            "items": {"type": "array", "items": {"type": "string"}},
                        },
                        "start_cell": {"type": "string", "description": "Start cell"},
                    },
                    "required": ["path", "sheet", "data", "start_cell"],
                },
            ),
            Tool(
                name="excel_set_column_width",
                description="Set column width",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "column": {"type": "string", "description": "Column letter"},
                        "width": {"type": "number", "description": "Width"},
                    },
                    "required": ["path", "sheet", "column", "width"],
                },
            ),
            Tool(
                name="excel_merge_cells",
                description="Merge cells",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                        "sheet": {"type": "string", "description": "Sheet name"},
                        "start_cell": {"type": "string", "description": "Start cell"},
                        "end_cell": {"type": "string", "description": "End cell"},
                    },
                    "required": ["path", "sheet", "start_cell", "end_cell"],
                },
            ),
            Tool(
                name="excel_list_sheets",
                description="List all sheets",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to workbook"},
                    },
                    "required": ["path"],
                },
            ),
            Tool(
                name="excel_to_pdf",
                description="Convert Excel to PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "input_path": {"type": "string", "description": "Excel path"},
                        "output_path": {"type": "string", "description": "PDF path"},
                    },
                    "required": ["input_path", "output_path"],
                },
            ),
        ]

    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        try:
            if tool_name == "excel_create":
                return self._create(arguments)
            elif tool_name == "excel_read":
                return self._read(arguments)
            elif tool_name == "excel_write_cell":
                return self._write_cell(arguments)
            elif tool_name == "excel_add_row":
                return self._add_row(arguments)
            elif tool_name == "excel_add_formula":
                return self._add_formula(arguments)
            elif tool_name == "excel_format_cell":
                return self._format_cell(arguments)
            elif tool_name == "excel_add_table":
                return self._add_table(arguments)
            elif tool_name == "excel_set_column_width":
                return self._set_column_width(arguments)
            elif tool_name == "excel_merge_cells":
                return self._merge_cells(arguments)
            elif tool_name == "excel_list_sheets":
                return self._list_sheets(arguments)
            elif tool_name == "excel_to_pdf":
                return self._excel_to_pdf(arguments)
            return self.error_result(f"Unknown tool: {tool_name}")
        except Exception as e:
            return self.error_result(str(e))

    def _get_sheet(self, wb, sheet_name: str):
        return wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    def _create(self, args: dict) -> list[TextContent]:
        wb = Workbook()
        if name := args.get("sheet_name"):
            wb.active.title = name
        wb.save(args["path"])
        return self.success_result(f"Workbook created: {args['path']}")

    def _read(self, args: dict) -> list[TextContent]:
        import json

        wb = load_workbook(args["path"], data_only=True)
        ws = self._get_sheet(wb, args.get("sheet", ""))
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([str(cell) if cell is not None else "" for cell in row])
        return [TextContent(type="text", text=json.dumps(data, indent=2))]

    def _write_cell(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        ws[args["cell"]] = args["value"]
        wb.save(args["path"])
        return self.success_result(f"Wrote {args['value']} to {args['cell']}")

    def _add_row(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        ws.append(args["data"])
        wb.save(args["path"])
        return self.success_result(f"Added row with {len(args['data'])} cells")

    def _add_formula(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        ws[args["cell"]] = args["formula"]
        wb.save(args["path"])
        return self.success_result(f"Added formula {args['formula']}")

    def _format_cell(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        cell = ws[args["cell"]]

        if args.get("bold"):
            cell.font = Font(bold=True)
        if size := args.get("font_size"):
            cell.font = Font(size=size)
        if fc := args.get("font_color"):
            rgb = hex_to_rgb(fc)
            cell.font = Font(color=rgb_to_hex(rgb))
        if bg := args.get("bg_color"):
            rgb = hex_to_rgb(bg)
            cell.fill = PatternFill(start_color=rgb_to_hex(rgb), fill_type="solid")

        wb.save(args["path"])
        return self.success_result(f"Formatted cell {args['cell']}")

    def _add_table(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        data = args["data"]
        start = ws[args["start_cell"]]

        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                cell = ws.cell(row=start.row + r_idx, column=start.column + c_idx)
                cell.value = val

        wb.save(args["path"])
        return self.success_result(f"Added table with {len(data)} rows")

    def _set_column_width(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        ws.column_dimensions[args["column"]].width = args["width"]
        wb.save(args["path"])
        return self.success_result(f"Set column {args['column']} width")

    def _merge_cells(self, args: dict) -> list[TextContent]:
        wb = load_workbook(args["path"])
        ws = self._get_sheet(wb, args["sheet"])
        ws.merge_cells(f"{args['start_cell']}:{args['end_cell']}")
        wb.save(args["path"])
        return self.success_result(f"Merged {args['start_cell']}:{args['end_cell']}")

    def _list_sheets(self, args: dict) -> list[TextContent]:
        import json

        wb = load_workbook(args["path"])
        return [TextContent(type="text", text=json.dumps(wb.sheetnames, indent=2))]

    def _excel_to_pdf(self, args: dict) -> list[TextContent]:
        return self.success_result(
            "Excel to PDF requires LibreOffice. Install and use: soffice --headless --convert-to pdf input.xlsx"
        )
