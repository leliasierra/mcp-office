# MCP Office

MCP server for Office document manipulation (Word, Excel, PowerPoint, PDF).

## Installation

```bash
uv sync
```

## Usage

```bash
uv run mcp-office
```

## Features

### Word Document Management
- Create new Word documents with metadata
- Extract text and analyze document structure
- Get document properties and statistics
- List documents in a directory
- Create copies of existing documents
- Merge multiple documents into one

### Word Content Creation
- Add headings with different levels and formatting
- Add paragraphs with style and formatting
- Create tables with custom data
- Add images with proportional scaling
- Insert page breaks
- Add bullet and numbered lists
- Add footnotes and endnotes
- Convert footnotes to endnotes

### Word Text Formatting
- Format specific text ranges (bold, italic, underline)
- Change text color and font properties
- Apply custom styles to text elements
- Find and replace text throughout the document
- Format individual table cells

### Word Table Formatting
- Format tables with borders and styles
- Create formatted header rows
- Apply cell shading and custom borders
- Alternating row colors
- Cell alignment (horizontal and vertical)
- Merge cells (horizontal, vertical, rectangular)
- Column width management (points, percentage, auto-fit)

### Word Document Protection
- Add password protection to documents
- Add editable regions

### Word Comments
- Extract all comments from a document
- Filter comments by author

### PDF Operations
- Read and extract content from PDF
- Get PDF metadata and information
- Extract images from PDF
- Split PDF into separate pages
- Merge multiple PDFs
- Rotate pages
- Compress PDF
- Add watermark
- Add page numbers
- Protect/unprotect PDF with password
- Extract text from specific pages
- Extract tables (experimental)

### Conversions
- Convert Word to PDF
- Convert PDF to Word