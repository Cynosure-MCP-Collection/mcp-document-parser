# @cynosure-mcp/document-parser

MCP server that extracts text from documents and returns it as Markdown.

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

| Tool                   | Description                                                                       |
| ---------------------- | --------------------------------------------------------------------------------- |
| `parse_document`       | Parse a document file and return its contents as Markdown text                    |
| `get_supported_formats`| Return the list of supported document formats                                     |

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
