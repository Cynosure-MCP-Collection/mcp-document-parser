#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { parseOffice, type OfficeContentNode, type OfficeAttachment } from 'officeparser';
import { z } from 'zod';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';

// ── Supported file extensions ──────────────────────────────────────────────────

const SUPPORTED_EXTENSIONS = new Set([
    '.docx', '.pptx', '.xlsx',
    '.odt', '.odp', '.ods',
    '.pdf', '.rtf'
]);

function isSupported(filename: string): boolean {
    const ext = path.extname(filename).toLowerCase();
    return SUPPORTED_EXTENSIONS.has(ext);
}

// ── AST → Markdown Conversion ──────────────────────────────────────────────────

function nodeToMarkdown(node: OfficeContentNode, depth = 0): string {
    const meta = node.metadata as Record<string, unknown> | undefined;
    switch (node.type) {
        case 'heading': {
            const level = Math.min((meta?.level as number) || 1, 6);
            return `${'#'.repeat(level)} ${node.text || ''}`;
        }

        case 'paragraph':
            return node.text || '';

        case 'list':
            return listItemToMarkdown(node);

        case 'table':
            return tableToMarkdown(node);

        case 'note':
            return `> **Note${meta?.noteType ? ` (${meta.noteType})` : ''}:** ${node.text || ''}`;

        case 'image':
            return `[Image: ${meta?.altText || meta?.attachmentName || 'embedded image'}]`;

        case 'chart':
            return `[Chart: ${meta?.attachmentName || 'embedded chart'}]`;

        default:
            return node.text || '';
    }
}

function listItemToMarkdown(node: OfficeContentNode): string {
    const meta = node.metadata as Record<string, unknown> | undefined;
    const indent = '  '.repeat((meta?.indentation as number) || 0);
    const isOrdered = meta?.listType === 'ordered';
    const idx = ((meta?.itemIndex as number) ?? 0) + 1;
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

    const meta = ast.metadata;
    if (meta) {
        const metaParts: string[] = [];
        if (meta.title) metaParts.push(`**Title:** ${meta.title}`);
        if (meta.author) metaParts.push(`**Author:** ${meta.author}`);
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
            .map(a => `**[OCR – ${a.name || 'image'}]:**\n${a.ocrText!.trim()}`);
        if (ocrTexts.length) {
            lines.push('');
            lines.push('---');
            lines.push('## Extracted Text from Images (OCR)');
            lines.push(...ocrTexts);
        }
    }

    return lines.join('\n\n');
}

// ── MCP Server ─────────────────────────────────────────────────────────────────

const server = new McpServer({
    name: 'document-parser',
    version: '1.0.0',
    description: 'MCP server that extracts text from documents and returns it as Markdown',
    icons: [{ src: 'https://raw.githubusercontent.com/andreasjhagen/Cynosure-MCPs/main/mcp-document-parser/icon.png', mimeType: 'image/png' }],
});

server.registerTool(
    'parse_document',
    {
        description: 'Parse a document file and return its contents as Markdown text. Supported formats: .docx, .pptx, .xlsx, .odt, .odp, .ods, .pdf, .rtf. OCR is available for image extraction if requested.',
        inputSchema: z.object({
            filePath: z.string().describe('Absolute path to the document file to parse.'),
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
                        text: `Error: Unsupported file type "${ext}". Supported extensions: ${[...SUPPORTED_EXTENSIONS].join(', ')}`,
                    },
                ],
                isError: true,
            };
        }

        let buffer: Buffer;
        try {
            buffer = await fs.readFile(resolved);
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return {
                content: [{ type: 'text', text: `Error reading file: ${message}` }],
                isError: true,
            };
        }

        try {
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
    'get_supported_formats',
    {
        description: 'Return the list of document formats supported by this parser.',
        inputSchema: z.object({}),
    },
    async () => ({
        content: [
            {
                type: 'text',
                text: `Supported document formats: ${[...SUPPORTED_EXTENSIONS].join(', ')}`,
            },
        ],
    })
);

// ── Start transport ────────────────────────────────────────────────────────────

const transport = new StdioServerTransport();
await server.connect(transport);
