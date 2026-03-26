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

### Excel Operations
- Create and read workbooks
- Write cells and add rows
- Add formulas
- Format cells (bold, color, background)
- Add tables and merge cells
- Set column widths
- List sheets

### PowerPoint Operations
- Create presentations
- Add slides with layouts
- Add titles and text
- Add images
- Add tables
- List slides

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
- Convert Excel to PDF (requires LibreOffice)
- Convert PowerPoint to PDF (requires LibreOffice)

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

## OpenCode CLI Usage

Run OpenCode with a direct message using `opencode run`:

```bash
# Create a Word document
opencode run "Create a new Word document at C:/reports/test.docx with title 'Quarterly Report' and add a paragraph 'Hello World'. use office"

# Read PDF content
opencode run "Extract all text from C:/documents/report.pdf. use office"

# Create Excel spreadsheet
opencode run "Create an Excel workbook at C:/data/sales.xlsx with sheet 'Sales' and add headers ['Product', 'Qty', 'Price']. use office"

# Add slides to PowerPoint
opencode run "Create a PowerPoint at C:/presentations/demo.pptx with title 'My Presentation' and add a title+content slide. use office"
```

## GitHub Copilot CLI Configuration

For GitHub Copilot CLI (`gh copilot`), add to your `~/.copilotrc` or configure via `gh copilot config`:

```json
{
  "mcp_servers": {
    "office": {
      "command": "python",
      "args": ["-c", "import asyncio;from mcp_office import main;asyncio.run(main())"],
      "env": {
        "PYTHONPATH": "path/to/mcp-office/src"
      }
    }
  }
}
```

Or set the environment variable:
```bash
export PYTHONPATH=path/to/mcp-office/src
gh copilot run --experimental-mcp "office" "Create a Word document..."
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
| `excel_create` | Create a new Excel workbook |
| `excel_read` | Read data from Excel |
| `excel_write_cell` | Write to a cell |
| `excel_add_row` | Add a row |
| `excel_add_formula` | Add a formula |
| `excel_format_cell` | Format a cell |
| `excel_add_table` | Add a table |
| `excel_set_column_width` | Set column width |
| `excel_merge_cells` | Merge cells |
| `excel_list_sheets` | List sheets |
| `ppt_create` | Create a presentation |
| `ppt_add_slide` | Add a slide |
| `ppt_add_title` | Add title to slide |
| `ppt_add_text` | Add text to slide |
| `ppt_add_image` | Add image to slide |
| `ppt_add_table` | Add table to slide |
| `ppt_list_slides` | List slides |
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
