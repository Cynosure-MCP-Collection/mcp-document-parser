#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { parseOffice, type OfficeAttachment, type OfficeContentNode } from 'officeparser';
import {
    Document,
    ExternalHyperlink,
    HeadingLevel,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType,
} from 'docx';
import { z } from 'zod';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';

// ── Supported file extensions ──────────────────────────────────────────────────

const SUPPORTED_READ_EXTENSIONS = new Set([
    '.docx', '.pptx', '.xlsx',
    '.odt', '.odp', '.ods',
    '.pdf', '.rtf'
]);

const DOCX_EXTENSIONS = new Set(['.docx']);

type DocxChild = Paragraph | Table;

function isSupported(filename: string): boolean {
    const ext = path.extname(filename).toLowerCase();
    return SUPPORTED_READ_EXTENSIONS.has(ext);
}

// ── AST → Markdown Conversion ──────────────────────────────────────────────────

function nodeToMarkdown(node: OfficeContentNode): string {
    const meta = node.metadata as Record<string, unknown> | undefined;
    switch (node.type) {
        case 'heading': {
            const level = Math.min(Number(meta?.level) || 1, 6);
            return `${'#'.repeat(level)} ${node.text || ''}`;
        }

        case 'paragraph':
            return node.text || '';

        case 'list':
            return listItemToMarkdown(node);

        case 'table':
            return tableToMarkdown(node);

        case 'note':
            return `> **Note${meta?.noteType ? ` (${String(meta.noteType)})` : ''}:** ${node.text || ''}`;

        case 'image':
            return `[Image: ${String(meta?.altText || meta?.attachmentName || 'embedded image')}]`;

        case 'chart':
            return `[Chart: ${String(meta?.attachmentName || 'embedded chart')}]`;

        default:
            return node.text || '';
    }
}

function listItemToMarkdown(node: OfficeContentNode): string {
    const meta = node.metadata as Record<string, unknown> | undefined;
    const indent = '  '.repeat(Math.max(0, Number(meta?.indentation) || 0));
    const isOrdered = meta?.listType === 'ordered';
    const idx = (Number(meta?.itemIndex) || 0) + 1;
    const bullet = isOrdered ? `${idx}.` : '-';
    return `${indent}${bullet} ${node.text || ''}`;
}

function tableToMarkdown(tableNode: OfficeContentNode): string {
    const rows = (tableNode.children || []).filter(r => r.type === 'row');
    if (!rows.length) return '';

    const tableRows: string[][] = rows.map(row =>
        (row.children || [])
            .filter(c => c.type === 'cell')
            .map(cell => (cell.text || '').replace(/\|/g, '\\|').replace(/\n/g, ' ').trim())
    );

    if (!tableRows.length || !tableRows[0].length) return '';

    const colCount = Math.max(...tableRows.map(r => r.length));
    const normalized = tableRows.map(r => {
        while (r.length < colCount) r.push('');
        return r;
    });

    const lines: string[] = [];
    lines.push('| ' + normalized[0].join(' | ') + ' |');
    lines.push('| ' + normalized[0].map(() => '---').join(' | ') + ' |');
    for (let i = 1; i < normalized.length; i++) {
        lines.push('| ' + normalized[i].join(' | ') + ' |');
    }

    return lines.join('\n');
}

/** Parse a document buffer and return structured Markdown text */
async function parseDocument(buffer: Buffer, filename: string, ocr = false, ocrLanguage = 'eng'): Promise<string> {
    const ast = await parseOffice(buffer, {
        outputErrorToConsole: false,
        extractAttachments: ocr,
        ocr,
        ocrLanguage,
    });

    const lines: string[] = [];

    const meta = ast.metadata as { title?: string; author?: string } | undefined;
    if (meta) {
        const metaParts: string[] = [];
        if (meta.title) metaParts.push(`**Title:** ${meta.title}`);
        if (meta.author && meta.author !== 'Un-named') metaParts.push(`**Author:** ${meta.author}`);
        if (metaParts.length) {
            lines.push(metaParts.join(' | '));
            lines.push('');
        }
    }

    for (const node of ast.content) {
        const md = nodeToMarkdown(node);
        if (md) lines.push(md);
    }

    if (ocr && ast.attachments?.length) {
        const ocrTexts = (ast.attachments as OfficeAttachment[])
            .filter(a => a.ocrText?.trim())
            .map(a => `**[OCR - ${a.name || 'image'}]:**\n${a.ocrText!.trim()}`);
        if (ocrTexts.length) {
            lines.push('');
            lines.push('---');
            lines.push('## Extracted Text from Images (OCR)');
            lines.push(...ocrTexts);
        }
    }

    return lines.join('\n\n');
}

// ── Markdown → DOCX Conversion ─────────────────────────────────────────────────

function splitInlineMarkdown(text: string): Array<TextRun | ExternalHyperlink> {
    const children: Array<TextRun | ExternalHyperlink> = [];
    const tokenRe = /(\*\*([^*]+)\*\*|\*([^*]+)\*|`([^`]+)`|\[([^\]]+)\]\(([^)]+)\))/g;
    let cursor = 0;
    let match: RegExpExecArray | null;

    while ((match = tokenRe.exec(text))) {
        if (match.index > cursor) children.push(new TextRun(text.slice(cursor, match.index)));
        if (match[2]) children.push(new TextRun({ text: match[2], bold: true }));
        else if (match[3]) children.push(new TextRun({ text: match[3], italics: true }));
        else if (match[4]) children.push(new TextRun({ text: match[4], font: 'Courier New' }));
        else if (match[5] && match[6]) {
            children.push(new ExternalHyperlink({
                link: match[6],
                children: [new TextRun({ text: match[5], style: 'Hyperlink' })],
            }));
        }
        cursor = match.index + match[0].length;
    }

    if (cursor < text.length) children.push(new TextRun(text.slice(cursor)));
    return children.length ? children : [new TextRun('')];
}

function paragraphFromMarkdown(line: string): Paragraph {
    const heading = line.match(/^(#{1,6})\s+(.+)$/);
    if (heading) {
        const levels = [
            HeadingLevel.HEADING_1,
            HeadingLevel.HEADING_2,
            HeadingLevel.HEADING_3,
            HeadingLevel.HEADING_4,
            HeadingLevel.HEADING_5,
            HeadingLevel.HEADING_6,
        ];
        return new Paragraph({
            text: heading[2],
            heading: levels[Math.min(heading[1].length, 6) - 1],
        });
    }

    const unordered = line.match(/^\s*[-*]\s+(.+)$/);
    if (unordered) {
        return new Paragraph({ children: splitInlineMarkdown(unordered[1]), bullet: { level: 0 } });
    }

    const ordered = line.match(/^\s*\d+\.\s+(.+)$/);
    if (ordered) {
        return new Paragraph({ children: splitInlineMarkdown(ordered[1]), numbering: { reference: 'default-numbering', level: 0 } });
    }

    const quote = line.match(/^>\s*(.+)$/);
    if (quote) {
        return new Paragraph({ children: splitInlineMarkdown(quote[1]), indent: { left: 360 } });
    }

    return new Paragraph({ children: splitInlineMarkdown(line) });
}

function tableFromMarkdown(lines: string[]): Table | null {
    if (lines.length < 2 || !/^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(lines[1])) return null;

    const rows = lines
        .filter((_, index) => index !== 1)
        .map(line => line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(cell => cell.trim().replace(/\\\|/g, '|')));

    if (!rows.length) return null;

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: rows.map(row => new TableRow({
            children: row.map(cell => new TableCell({
                children: [new Paragraph({ children: splitInlineMarkdown(cell) })],
            })),
        })),
    });
}

function markdownToDocxChildren(markdown: string): DocxChild[] {
    const children: DocxChild[] = [];
    const lines = markdown.replace(/\r\n/g, '\n').split('\n');

    for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i].trimEnd();
        if (!line.trim()) {
            children.push(new Paragraph(''));
            continue;
        }

        if (line.trimStart().startsWith('|')) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].trimStart().startsWith('|')) {
                tableLines.push(lines[i]);
                i += 1;
            }
            i -= 1;
            const table = tableFromMarkdown(tableLines);
            if (table) {
                children.push(table);
                continue;
            }
            children.push(...tableLines.map(paragraphFromMarkdown));
            continue;
        }

        children.push(paragraphFromMarkdown(line));
    }

    return children.length ? children : [new Paragraph('')];
}

async function writeDocx(markdown: string, outputPath: string, overwrite: boolean): Promise<string> {
    let resolved = path.resolve(outputPath);
    if (!DOCX_EXTENSIONS.has(path.extname(resolved).toLowerCase())) {
        resolved += '.docx'
    }

    await fs.mkdir(path.dirname(resolved), { recursive: true });
    if (!overwrite) {
        try {
            await fs.access(resolved);
            throw new Error(`File already exists: ${resolved}`);
        } catch (err) {
            if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
        }
    }

    const doc = new Document({
        numbering: {
            config: [{
                reference: 'default-numbering',
                levels: [{
                    level: 0,
                    format: 'decimal',
                    text: '%1.',
                    alignment: 'left',
                }],
            }],
        },
        sections: [{ children: markdownToDocxChildren(markdown) }],
    });

    await fs.writeFile(resolved, await Packer.toBuffer(doc));
    return resolved;
}

function applyDocumentEdits(markdown: string, edits: {
    prependMarkdown?: string;
    appendMarkdown?: string;
    replacements?: { search: string; replace: string }[];
}): string {
    let updated = markdown;
    if (edits.replacements?.length) {
        for (const replacement of edits.replacements) {
            if (!replacement.search) continue;
            updated = updated.split(replacement.search).join(replacement.replace);
        }
    }
    if (edits.prependMarkdown?.trim()) updated = `${edits.prependMarkdown.trim()}\n\n${updated}`.trim();
    if (edits.appendMarkdown?.trim()) updated = `${updated.trim()}\n\n${edits.appendMarkdown.trim()}`.trim();
    return updated;
}

function stripParserMetadata(markdown: string): string {
    return markdown
        .replace(/^\*\*(?:Title|Author):\*\*[^\n]*(?:\s+\|\s+\*\*(?:Title|Author):\*\*[^\n]*)*\n{2,}/, '')
        .trim();
}

// ── MCP Server ─────────────────────────────────────────────────────────────────

const server = new McpServer({
    name: 'document-reader-writer',
    version: '1.1.0',
    description: 'Read documents as Markdown and create or edit DOCX files.',
    icons: [{ src: 'https://raw.githubusercontent.com/andreasjhagen/Cynosure-MCPs/main/mcp-document-parser/icon.png', mimeType: 'image/png' }],
});

server.registerTool(
    'parse_document',
    {
        annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
        description: 'Parse a document file and return its contents as Markdown text. Supported formats: .docx, .pptx, .xlsx, .odt, .odp, .ods, .pdf, .rtf. OCR is available for image extraction if requested.',
        inputSchema: z.object({
            filePath: z.string().describe('Absolute or relative path to the document file to parse.'),
            ocr: z.boolean().optional().default(false).describe('Enable OCR on embedded images within the document.'),
            ocrLanguage: z.string().optional().default('eng').describe('OCR language code (e.g., "eng", "deu", "fra"). Only used when OCR is enabled.'),
        }),
    },
    async ({ filePath, ocr, ocrLanguage }) => {
        const resolved = path.resolve(filePath);
        const ext = path.extname(resolved).toLowerCase();

        if (!isSupported(resolved)) {
            return {
                content: [
                    {
                        type: 'text',
                        text: `Error: Unsupported file type "${ext}". Supported extensions: ${[...SUPPORTED_READ_EXTENSIONS].join(', ')}`,
                    },
                ],
                isError: true,
            };
        }

        try {
            const buffer = await fs.readFile(resolved);
            const markdown = await parseDocument(buffer, path.basename(resolved), ocr, ocrLanguage);
            return {
                content: [
                    {
                        type: 'text',
                        text: `[Parsed document: ${path.basename(resolved)}]\n\n${markdown}`,
                    },
                ],
            };
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return {
                content: [{ type: 'text', text: `Error parsing document: ${message}` }],
                isError: true,
            };
        }
    }
);

server.registerTool(
    'create_docx',
    {
        annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
        description: 'Create and save a DOCX file from Markdown content. Returns the saved local file path so clients can show it as a downloadable artifact.',
        inputSchema: z.object({
            markdown: z.string().describe('Markdown content to write into the DOCX document. Supports headings, paragraphs, simple lists, links, inline bold/italic/code, and Markdown tables.'),
            outputPath: z.string().describe('Where to save the DOCX file. Relative paths resolve from the MCP process working directory. The .docx extension is appended automatically if not provided.'),
            overwrite: z.boolean().optional().default(false).describe('Overwrite the output file if it already exists.'),
        }),
    },
    async ({ markdown, outputPath, overwrite }) => {
        try {
            const savedPath = await writeDocx(markdown, outputPath, overwrite);
            return {
                content: [{
                    type: 'text',
                    text: `Created DOCX document: ${savedPath}\n\nArtifact: [${path.basename(savedPath)}](${savedPath})`,
                }],
            };
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return { content: [{ type: 'text', text: `Error creating DOCX document: ${message}` }], isError: true };
        }
    }
);

server.registerTool(
    'edit_docx',
    {
        annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
        description: 'Edit an existing DOCX by parsing it to Markdown, applying text replacements and/or prepend/append Markdown, then saving a normalized DOCX copy. Existing complex Word formatting is not preserved.',
        inputSchema: z.object({
            filePath: z.string().describe('Path to the existing DOCX file to edit.'),
            outputPath: z.string().describe('Where to save the edited DOCX file. The .docx extension is appended automatically if not provided. Use a different path unless overwrite is true.'),
            replacements: z.array(z.object({
                search: z.string().describe('Exact text to find in the parsed Markdown.'),
                replace: z.string().describe('Replacement text.'),
            })).optional().default([]).describe('Exact text replacements to apply before saving.'),
            prependMarkdown: z.string().optional().default('').describe('Markdown to add to the beginning of the document.'),
            appendMarkdown: z.string().optional().default('').describe('Markdown to add to the end of the document.'),
            overwrite: z.boolean().optional().default(false).describe('Overwrite the output file if it already exists.'),
        }),
    },
    async ({ filePath, outputPath, replacements, prependMarkdown, appendMarkdown, overwrite }) => {
        const resolved = path.resolve(filePath);
        if (path.extname(resolved).toLowerCase() !== '.docx') {
            return { content: [{ type: 'text', text: 'Error editing DOCX document: input file must end with .docx' }], isError: true };
        }

        try {
            const buffer = await fs.readFile(resolved);
            const markdown = stripParserMetadata(await parseDocument(buffer, path.basename(resolved)));
            const updated = applyDocumentEdits(markdown, { replacements, prependMarkdown, appendMarkdown });
            const savedPath = await writeDocx(updated, outputPath, overwrite);
            return {
                content: [{
                    type: 'text',
                    text: `Edited DOCX document: ${savedPath}\n\nArtifact: [${path.basename(savedPath)}](${savedPath})`,
                }],
            };
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return { content: [{ type: 'text', text: `Error editing DOCX document: ${message}` }], isError: true };
        }
    }
);

server.registerTool(
    'get_supported_formats',
    {
        annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
        description: 'Return the list of document formats supported by this reader/writer.',
        inputSchema: z.object({}),
    },
    async () => ({
        content: [
            {
                type: 'text',
                text: [
                    `Readable document formats: ${[...SUPPORTED_READ_EXTENSIONS].join(', ')}`,
                    'Writable document formats: .docx',
                ].join('\n'),
            },
        ],
    })
);

// ── Start transport ────────────────────────────────────────────────────────────

const transport = new StdioServerTransport();
await server.connect(transport);
