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

### PowerPoint Operations (ppt-master Engine)
- Generate native editable PPTX from SVG files with full DrawingML shapes
- Support 8 canvas formats: 16:9, 4:3, A4, social media, banners
- Slide transitions, entrance animations, auto-advance (kiosk mode)
- Automatic SVG post-processing (icon embedding, image alignment, text flattening)
- Source document conversion to Markdown (PDF, DOCX, PPTX, XLSX, URL, and more)

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

### Source Document Conversion
- `ppt_convert_source` converts PDF, DOCX, PPTX, XLSX, HTML, EPUB, URLs, and more to Markdown
- Powered by ppt-master's source_to_md converters with Pandoc fallback for legacy formats

## Installation

```bash
uv sync
```

### Dependencies

- `python-docx` - Word document creation and manipulation
- `pypdf` - PDF reading and manipulation
- `pymupdf` - Advanced PDF operations
- `docx2pdf` - Word to PDF conversion
- `mcp` - MCP protocol SDK
- `python-pptx` - PowerPoint presentation creation
- `openpyxl` - Excel workbook reading/writing
- `xlsxwriter` - Excel workbook creation with advanced formatting
- `pandas` - Data analysis and Excel integration
- `reportlab` - PDF creation and text/watermark operations
- `img2pdf` - Image to PDF conversion
- `pytesseract` - OCR text extraction from images
- `Pillow` - Image processing (ppt-master SVG finalization)
- `svglib` - SVG parsing utilities (ppt-master SVG finalization)

### ppt-master Dependencies (optionally installed)

The ppt-master source converters require additional packages per converter:

- `PyMuPDF` — PDF to Markdown
- `mammoth`, `beautifulsoup4`, `markdownify`, `ebooklib`, `nbconvert` — Document to Markdown
- `python-pptx` — PowerPoint to Markdown
- `openpyxl`, `pandas` — Excel to Markdown
- `requests`, `beautifulsoup4` — Web to Markdown
- `curl_cffi` — TLS fingerprint bypass for WeChat/some Chinese sites
- `pandoc` (system executable) — Legacy format fallback (.doc, .odt, .rtf, .tex, etc.)

Install them as needed:
```bash
pip install PyMuPDF mammoth beautifulsoup4 markdownify ebooklib nbconvert requests curl_cffi
```

## OpenCode Configuration

Add to your `opencode.json`:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "office": {
      "type": "local",
      "command": ["C:\\Dev\\htdocs-uv\\mcp-project\\.venv\\Scripts\\python.exe", "-c", "import asyncio;from mcp_office import main;asyncio.run(main())"],
      "enabled": true,
      "environment": {
        "PYTHONPATH": "C:\\Dev\\htdocs-uv\\mcp-project\\src"
      },
      "timeout": 30000
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

# Generate PowerPoint from SVGs
opencode run "Generate a PPTX from SVGs in ./slides at ./output.pptx with format ppt169 and fade transitions. use office"

# Convert source document to Markdown
opencode run "Convert report.pdf to Markdown. use office"
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
| `ppt_generate` | Generate PPTX from SVG files (ppt-master engine) |
| `ppt_convert_source` | Convert source document (PDF/DOCX/PPTX/URL) to Markdown |
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
