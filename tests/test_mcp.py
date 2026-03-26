import pytest
import json
from pathlib import Path
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from mcp_office import list_tools, call_tool


@pytest.fixture
def temp_doc(tmp_path):
    doc_path = tmp_path / "test.docx"
    doc = Document()
    doc.add_paragraph("Test paragraph")
    doc.save(str(doc_path))
    return str(doc_path)


@pytest.fixture
def temp_xlsx(tmp_path):
    xlsx_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.active.append(["Header1", "Header2"])
    wb.save(str(xlsx_path))
    return str(xlsx_path)


@pytest.fixture
def temp_pptx(tmp_path):
    pptx_path = tmp_path / "test.pptx"
    prs = Presentation()
    prs.save(str(pptx_path))
    return str(pptx_path)


@pytest.fixture
def temp_pdf(tmp_path):
    pdf_path = tmp_path / "test.pdf"
    return str(pdf_path)


@pytest.mark.asyncio
async def test_list_tools():
    tools = await list_tools()
    assert len(tools) > 0
    tool_names = [t.name for t in tools]
    assert "create_document" in tool_names
    assert "open_document" in tool_names
    assert "pdf_read" in tool_names
    assert "excel_create" in tool_names
    assert "ppt_create" in tool_names


@pytest.mark.asyncio
async def test_create_document(tmp_path):
    path = str(tmp_path / "new_doc.docx")
    result = await call_tool("create_document", {"path": path})
    assert "created" in result[0].text.lower()
    assert Path(path).exists()


@pytest.mark.asyncio
async def test_open_document(temp_doc):
    result = await call_tool("open_document", {"path": temp_doc, "extract_text": True})
    data = json.loads(result[0].text)
    assert "paragraphs" in data
    assert data["paragraphs"] >= 1


@pytest.mark.asyncio
async def test_list_documents(tmp_path):
    doc_path = tmp_path / "test.docx"
    doc = Document()
    doc.add_paragraph("Test")
    doc.save(str(doc_path))

    result = await call_tool("list_documents", {"directory": str(tmp_path)})
    files = json.loads(result[0].text)
    assert len(files) >= 1
    assert any("test.docx" in f for f in files)


@pytest.mark.asyncio
async def test_get_document_properties(temp_doc):
    result = await call_tool("get_document_properties", {"path": temp_doc})
    data = json.loads(result[0].text)
    assert "paragraphs" in data
    assert "tables" in data


@pytest.mark.asyncio
async def test_copy_document(temp_doc, tmp_path):
    dest = str(tmp_path / "copy.docx")
    result = await call_tool("copy_document", {"source": temp_doc, "destination": dest})
    assert "copied" in result[0].text.lower()
    assert Path(dest).exists()


@pytest.mark.asyncio
async def test_add_paragraph(temp_doc):
    result = await call_tool(
        "add_paragraph", {"path": temp_doc, "text": "New paragraph", "bold": True}
    )
    assert "paragraph" in result[0].text.lower()


@pytest.mark.asyncio
async def test_add_heading(temp_doc):
    result = await call_tool(
        "add_heading",
        {"path": temp_doc, "text": "Test Heading", "level": 1, "bold": True},
    )
    assert "heading" in result[0].text.lower()


@pytest.mark.asyncio
async def test_find_replace(temp_doc):
    await call_tool("add_paragraph", {"path": temp_doc, "text": "Hello World"})
    result = await call_tool(
        "find_replace", {"path": temp_doc, "find": "World", "replace": "Python"}
    )
    assert "replaced" in result[0].text.lower()


@pytest.mark.asyncio
async def test_add_table(temp_doc):
    data = [["Header1", "Header2"], ["Row1Col1", "Row1Col2"]]
    result = await call_tool(
        "add_table", {"path": temp_doc, "data": data, "header_row": True}
    )
    assert "table" in result[0].text.lower()


@pytest.mark.asyncio
async def test_add_page_break(temp_doc):
    result = await call_tool("add_page_break", {"path": temp_doc})
    assert "page" in result[0].text.lower()


@pytest.mark.asyncio
async def test_bullet_list(temp_doc):
    result = await call_tool(
        "add_bullet_list", {"path": temp_doc, "items": ["Item 1", "Item 2", "Item 3"]}
    )
    assert "list" in result[0].text.lower()


@pytest.mark.asyncio
async def test_invalid_tool():
    result = await call_tool("nonexistent_tool", {})
    assert "unknown tool" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_create(tmp_path):
    path = str(tmp_path / "new_book.xlsx")
    result = await call_tool("excel_create", {"path": path})
    assert "created" in result[0].text.lower()
    assert Path(path).exists()


@pytest.mark.asyncio
async def test_excel_read(temp_xlsx):
    result = await call_tool("excel_read", {"path": temp_xlsx})
    data = json.loads(result[0].text)
    assert len(data) > 0
    assert "Header1" in data[0]


@pytest.mark.asyncio
async def test_excel_write_cell(temp_xlsx):
    result = await call_tool(
        "excel_write_cell",
        {"path": temp_xlsx, "sheet": "Sheet", "cell": "A2", "value": "TestValue"},
    )
    assert "wrote" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_add_row(temp_xlsx):
    result = await call_tool(
        "excel_add_row",
        {"path": temp_xlsx, "sheet": "Sheet", "data": ["Data1", "Data2"]},
    )
    assert "added" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_add_formula(temp_xlsx):
    result = await call_tool(
        "excel_add_formula",
        {"path": temp_xlsx, "sheet": "Sheet", "cell": "C1", "formula": "=SUM(A1:B1)"},
    )
    assert "formula" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_format_cell(temp_xlsx):
    result = await call_tool(
        "excel_format_cell",
        {
            "path": temp_xlsx,
            "sheet": "Sheet",
            "cell": "A1",
            "bold": True,
            "font_size": 14,
        },
    )
    assert "formatted" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_add_table(temp_xlsx):
    result = await call_tool(
        "excel_add_table",
        {
            "path": temp_xlsx,
            "sheet": "Sheet",
            "data": [["A", "B"], ["1", "2"]],
            "start_cell": "E1",
        },
    )
    assert "table" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_set_column_width(temp_xlsx):
    result = await call_tool(
        "excel_set_column_width",
        {"path": temp_xlsx, "sheet": "Sheet", "column": "A", "width": 20},
    )
    assert "width" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_merge_cells(temp_xlsx):
    result = await call_tool(
        "excel_merge_cells",
        {"path": temp_xlsx, "sheet": "Sheet", "start_cell": "A1", "end_cell": "B1"},
    )
    assert "merged" in result[0].text.lower()


@pytest.mark.asyncio
async def test_excel_list_sheets(temp_xlsx):
    result = await call_tool("excel_list_sheets", {"path": temp_xlsx})
    sheets = json.loads(result[0].text)
    assert len(sheets) >= 1


@pytest.mark.asyncio
async def test_ppt_create(tmp_path):
    path = str(tmp_path / "new_presentation.pptx")
    result = await call_tool("ppt_create", {"path": path, "title": "My Presentation"})
    assert "created" in result[0].text.lower()
    assert Path(path).exists()


@pytest.mark.asyncio
async def test_ppt_add_slide(temp_pptx):
    result = await call_tool("ppt_add_slide", {"path": temp_pptx, "layout": "title"})
    assert "added" in result[0].text.lower()


@pytest.mark.asyncio
async def test_ppt_add_title(temp_pptx):
    await call_tool("ppt_add_slide", {"path": temp_pptx, "layout": "title_content"})
    result = await call_tool(
        "ppt_add_title",
        {"path": temp_pptx, "slide_index": 0, "title": "Slide Title", "font_size": 32},
    )
    assert "title" in result[0].text.lower()


@pytest.mark.asyncio
async def test_ppt_add_text(temp_pptx):
    await call_tool("ppt_add_slide", {"path": temp_pptx, "layout": "blank"})
    result = await call_tool(
        "ppt_add_text",
        {
            "path": temp_pptx,
            "slide_index": 0,
            "text": "Sample text",
            "left": 1,
            "top": 2,
            "width": 5,
            "height": 1,
        },
    )
    assert "text" in result[0].text.lower()


@pytest.mark.asyncio
async def test_ppt_list_slides(temp_pptx):
    result = await call_tool("ppt_list_slides", {"path": temp_pptx})
    slides = json.loads(result[0].text)
    assert isinstance(slides, list)
