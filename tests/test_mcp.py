import pytest
import json
import asyncio
from pathlib import Path
from docx import Document
from mcp_office import server, list_tools, call_tool


@pytest.fixture
def temp_doc(tmp_path):
    doc_path = tmp_path / "test.docx"
    doc = Document()
    doc.add_paragraph("Test paragraph")
    doc.save(str(doc_path))
    return str(doc_path)


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
