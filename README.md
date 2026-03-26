# MCP Office

MCP server for Office document manipulation (Word, Excel, PowerPoint, PDF).

## Features

### Word Document Management
- Create, open, and analyze Word documents
- Get document properties and statistics
- List, copy, and merge documents

### Word Content Creation
- Add headings with formatting
- Add paragraphs with style and formatting
- Create tables with custom data
- Add images with proportional scaling
- Insert page breaks
- Add bullet and numbered lists
- Add footnotes and endnotes

### Word Text Formatting
- Format text ranges (bold, italic, underline)
- Change text color and font properties
- Find and replace text

### Word Table Formatting
- Format tables with borders and styles
- Create formatted header rows
- Merge cells
- Column width management

### PDF Operations
- Read and extract content from PDF
- Get PDF metadata
- Extract images from PDF
- Split and merge PDFs
- Rotate pages
- Compress PDF
- Add watermark and page numbers
- Protect/unprotect PDF with password

### Conversions
- Convert Word to PDF
- Convert PDF to Word

## Installation

```bash
uv sync
```

## OpenCode Configuration

Add to your `opencode.json`:

```json
{
  "mcp": {
    "office": {
      "type": "local",
      "command": ["python", "-c", "import asyncio;from mcp_office import main;asyncio.run(main())"],
      "enabled": true,
      "environment": {
        "PYTHONPATH": "path/to/mcp-office/src"
      }
    }
  }
}
```

Then use it in your prompts:
```
Create a new Word document with the quarterly report. use office
Extract all text from document.pdf. use office
```

## Available Tools

| Tool | Description |
|------|-------------|
| `create_document` | Create a new Word document |
| `open_document` | Open and analyze a Word document |
| `get_document_properties` | Get document properties and statistics |
| `list_documents` | List Word documents in a directory |
| `copy_document` | Create a copy of a document |
| `merge_documents` | Merge multiple documents |
| `add_heading` | Add a heading |
| `add_paragraph` | Add a paragraph |
| `add_table` | Create a table |
| `add_image` | Add an image |
| `add_page_break` | Insert a page break |
| `add_bullet_list` | Add a bullet/numbered list |
| `find_replace` | Find and replace text |
| `format_table_cell` | Format table cells |
| `pdf_read` | Read and extract PDF content |
| `pdf_get_info` | Get PDF metadata |
| `pdf_split` | Split PDF into pages |
| `pdf_merge` | Merge multiple PDFs |
| `pdf_rotate` | Rotate PDF pages |
| `pdf_compress` | Compress PDF |
| `pdf_add_watermark` | Add watermark to PDF |
| `pdf_protect` | Protect PDF with password |
| `word_to_pdf` | Convert Word to PDF |
| `pdf_to_word` | Convert PDF to Word |

## Development

```bash
# Install dependencies
uv sync

# Run linting
uv run ruff check src/

# Run tests
uv run pytest tests/
```

## License

MIT
