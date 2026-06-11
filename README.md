# @cynosure-mcp/document-parser

MCP server that reads documents as Markdown and creates or edits DOCX files.

## Installation

```bash
npx @cynosure-mcp/document-parser
```

Or install globally:

```bash
npm install -g @cynosure-mcp/document-parser
document-parser
```

## Tools

| Tool                    | Description                                                                       |
| ----------------------- | --------------------------------------------------------------------------------- |
| `parse_document`        | Parse a document file and return its contents as Markdown text                    |
| `create_docx`           | Create and save a DOCX file from Markdown content                                |
| `edit_docx`             | Edit an existing DOCX by replacing text and/or prepending/appending Markdown      |
| `get_supported_formats` | Return the list of supported reader/writer formats                               |

### Supported Formats

| Format | Extension |
| ------ | --------- |
| Word   | .docx     |
| PowerPoint | .pptx |
| Excel  | .xlsx     |
| PDF    | .pdf      |
| Rich Text | .rtf   |
| OpenDocument Text | .odt |
| OpenDocument Presentation | .odp |
| OpenDocument Spreadsheet | .ods |

### Writable Format

| Format | Extension |
| ------ | --------- |
| Word   | .docx     |

`create_docx` supports headings, paragraphs, simple ordered/unordered lists, links, inline bold/italic/code, and Markdown tables.

`edit_docx` parses an existing DOCX to Markdown, applies exact text replacements plus optional prepend/append Markdown, then saves a normalized DOCX copy. Complex Word formatting from the original document is not preserved.

## Configuration

No configuration required.

### MCP Config

```json
{
  "mcpServers": {
    "document-parser": {
      "command": "npx",
      "args": ["@cynosure-mcp/document-parser"]
    }
  }
}
```

## License

MIT
