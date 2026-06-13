// ============================================================
// @openthesis/docx-renderer — Template-Driven DOCX Renderer
// ============================================================
// Converts ThesisDocument content into a standards-compliant
// .docx file, using a parsed DocumentTemplate for all formatting.
// ============================================================

import {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, PageBreak,
  AlignmentType,
  BorderStyle, WidthType, ShadingType,
  convertMillimetersToTwip, Packer,
} from 'docx';
import type {
  DocumentTemplate,
  ThesisDocument,
  ContentBlock,
  ParagraphStyle,
  BlockType,
  LegacyDocumentJSON,
} from '@openthesis/document-schema';
import { writeFileSync } from 'fs';

// ── Renderer Config ───────────────────────────────────────

export interface RenderOptions {
  /** Path to write the output .docx */
  outputPath: string;
  /** Parsed template (from template-parser) */
  template: DocumentTemplate;
  /** Content to render */
  document: ThesisDocument;
  /** Header text (defaults to template.meta.organization + '学位论文') */
  headerText?: string;
  /** Show page numbers in footer */
  showPageNumbers?: boolean;
}

// ── Default Fallback Styles (when template lacks a role) ──

const DEFAULT_STYLES: Record<string, ParagraphStyle> = {
  heading1: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 30, bold: true },
    paragraph: { alignment: 'center', spaceBefore: 340, spaceAfter: 340 },
    lineSpacing: 312,
  },
  heading2: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 28, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 260, spaceAfter: 260 },
    lineSpacing: 312,
  },
  heading3: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 26, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 200, spaceAfter: 200 },
    lineSpacing: 312,
  },
  paragraph: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24 },
    paragraph: { alignment: 'justified', firstLineIndent: convertMillimetersToTwip(7.4) },
    lineSpacing: 312,
  },
  paragraph_no_indent: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24 },
    paragraph: { alignment: 'left' },
    lineSpacing: 312,
  },
  centered_text: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 30, bold: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 312,
  },
  equation: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24, italic: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 312,
  },
};

// ── Style Resolver ────────────────────────────────────────

/**
 * Resolve the ParagraphStyle for a given block type.
 * Priority: template.styles → defaults
 */
function resolveStyle(template: DocumentTemplate, blockType: BlockType): ParagraphStyle {
  // Look up via role mapping
  const styleId = Object.entries(template.styleRoles).find(
    ([, role]) => role === blockType,
  )?.[0];

  if (styleId && template.styles[styleId]) {
    return template.styles[styleId];
  }

  // Fallback
  return DEFAULT_STYLES[blockType] || DEFAULT_STYLES['paragraph'];
}

// ── Block Renderers ───────────────────────────────────────

function renderHeading(
  template: DocumentTemplate,
  block: ContentBlock & { text: string; number?: string },
): Paragraph {
  const style = resolveStyle(template, block.type);
  const numberPrefix = block.number ? `${block.number}  ` : '';
  return new Paragraph({
    children: [
      new TextRun({
        text: numberPrefix + block.text,
        bold: style.font.bold,
        size: style.font.size,
        font: {
          name: style.font.name,
          eastAsia: style.font.eastAsia,
        },
        italics: style.font.italic,
        color: style.font.color,
      }),
    ],
    alignment: mapAlignment(style.paragraph.alignment),
    spacing: {
      before: style.paragraph.spaceBefore,
      after: style.paragraph.spaceAfter,
      line: style.lineSpacing,
    },
  });
}

function renderParagraph(
  template: DocumentTemplate,
  block: ContentBlock & { text: string },
): Paragraph {
  const style = resolveStyle(template, block.type);
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size: style.font.size,
        font: {
          name: style.font.name,
          eastAsia: style.font.eastAsia,
        },
        bold: style.font.bold,
        italics: style.font.italic,
        color: style.font.color,
      }),
    ],
    alignment: mapAlignment(style.paragraph.alignment),
    indent: style.paragraph.firstLineIndent
      ? { firstLine: style.paragraph.firstLineIndent }
      : undefined,
    spacing: {
      before: style.paragraph.spaceBefore ?? 12,
      after: style.paragraph.spaceAfter ?? 12,
      line: style.lineSpacing,
    },
  });
}

function renderCenteredText(
  template: DocumentTemplate,
  block: ContentBlock & { text: string; font_size_pt?: number; bold?: boolean },
): Paragraph {
  const style = resolveStyle(template, 'centered_text');
  const actualSize = block.font_size_pt
    ? block.font_size_pt * 2  // convert pt → half-pt
    : style.font.size;
  const actualBold = block.bold !== undefined ? block.bold : style.font.bold;
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size: actualSize,
        bold: actualBold,
        font: {
          name: style.font.name,
          eastAsia: style.font.eastAsia,
        },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { line: style.lineSpacing },
  });
}

function renderEquation(
  _template: DocumentTemplate,
  block: ContentBlock & { latex: string; number?: string },
): Paragraph {
  // For now, render as plain text italic (V3 will add OMML)
  // This matches the existing generate_docx.mjs behavior
  const style = _template.styles['_defaults'] || DEFAULT_STYLES['equation'];
  const equationText = block.latex;
  const numberSuffix = block.number ? `    (${block.number})` : '';

  return new Paragraph({
    children: [
      new TextRun({
        text: equationText + numberSuffix,
        size: 24,
        font: 'Times New Roman',
        italics: true,
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: style.lineSpacing || 312 },
  });
}

function renderTable(
  template: DocumentTemplate,
  block: ContentBlock & {
    caption: string;
    headers: string[];
    data: string[][];
    showGridlines?: boolean;
    headerShading?: boolean;
  },
): [Paragraph, Table, Paragraph] {
  const colCount = block.headers.length;
  const totalWidth = 8504; // A4 printable width in DXA
  const colWidth = Math.floor(totalWidth / colCount);

  const showGrid = block.showGridlines !== false;
  const showShading = block.headerShading !== false;

  const BO = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
  const BI = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
  const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

  // Caption
  const captionPara = new Paragraph({
    children: [
      new TextRun({
        text: block.caption,
        bold: true,
        size: 21,
        font: { name: 'Times New Roman', eastAsia: '黑体' },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60, line: 312 },
  });

  // Header row
  const headerRow = new TableRow({
    tableHeader: true,
    children: block.headers.map((h, ci) =>
      new TableCell({
        width: { size: colWidth, type: WidthType.DXA },
        borders: showGrid ? {
          top: BO,
          bottom: ci === colCount - 1 ? BO : BI,
          left: ci === 0 ? BO : BI,
          right: ci === colCount - 1 ? BO : BI,
        } : {
          top: NO, bottom: NO, left: NO, right: NO,
        },
        shading: showShading ? { fill: 'D9D9D9', type: ShadingType.SOLID, color: 'auto' } : undefined,
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: h,
                bold: true,
                size: 21,
                font: { name: 'Times New Roman', eastAsia: '黑体' },
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 40 },
          }),
        ],
      }),
    ),
  });

  // Data rows
  const dataRows = block.data.map((row, ri) => {
    const isLast = ri === block.data.length - 1;
    return new TableRow({
      children: row.map((cell, ci) =>
        new TableCell({
          width: { size: colWidth, type: WidthType.DXA },
          borders: showGrid ? {
            top: BI,
            bottom: isLast ? BO : BI,
            left: ci === 0 ? BO : BI,
            right: ci === colCount - 1 ? BO : BI,
          } : {
            top: NO, bottom: NO, left: NO, right: NO,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: cell,
                  size: 21,
                  font: { name: 'Times New Roman', eastAsia: '宋体' },
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 20, after: 20 },
            }),
          ],
        }),
      ),
    });
  });

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: showGrid ? {
      top: BO, bottom: BO, left: BO, right: BO,
      insideHorizontal: BI, insideVertical: BI,
    } : {
      top: NO, bottom: NO, left: NO, right: NO,
      insideHorizontal: NO, insideVertical: NO,
    },
    rows: [headerRow, ...dataRows],
  });

  // Spacer after table
  const spacer = new Paragraph({ spacing: { line: 312 }, children: [] });

  return [captionPara, table, spacer];
}

function renderFigure(
  _template: DocumentTemplate,
  block: ContentBlock & { caption: string; path: string },
): Paragraph {
  // For now, insert a placeholder paragraph (full image support requires
  // reading the image file and embedding it — planned for V2)
  return new Paragraph({
    children: [
      new TextRun({
        text: `[图: ${block.caption}]`,
        size: 21,
        font: { name: 'Times New Roman', eastAsia: '宋体' },
        italics: true,
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: 312 },
  });
}

function renderSpacer(_template: DocumentTemplate, lines: number): Paragraph[] {
  return Array.from({ length: lines }, () =>
    new Paragraph({
      spacing: { line: 312 },
      children: [],
    }),
  );
}

function renderPageBreak(): Paragraph {
  return new Paragraph({ children: [new PageBreak()] });
}

function renderHorizontalRule(): Paragraph {
  // TODO: implement with a border-bottom hack or a drawing element
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
  });
}

// ── Official Document Block Renderers ──────────────────────

function renderRedHeader(
  block: ContentBlock & { text: string; font_size_pt?: number },
): Paragraph {
  const size = (block.font_size_pt || 22) * 2; // pt → half-pt
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size,
        bold: true,
        color: 'FF0000',
        font: { name: 'Times New Roman', eastAsia: '方正小标宋简体' },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: 312 },
  });
}

function renderDocumentNumber(
  block: ContentBlock & { text: string },
): Paragraph {
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size: 28, // 14pt
        font: { name: 'Times New Roman', eastAsia: '仿宋' },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60, line: 312 },
  });
}

function renderRecipientLine(
  block: ContentBlock & { text: string; recipientType: 'primary' | 'cc' },
): Paragraph {
  const prefix = block.recipientType === 'primary' ? '' : '';
  return new Paragraph({
    children: [
      new TextRun({
        text: prefix + block.text,
        size: 28, // 14pt
        font: { name: 'Times New Roman', eastAsia: '仿宋' },
      }),
    ],
    alignment: AlignmentType.LEFT,
    spacing: { before: 60, after: 60, line: 312 },
  });
}

function renderSignatureBlock(
  block: ContentBlock & { authority: string; date: string },
): [Paragraph, Paragraph] {
  return [
    new Paragraph({
      children: [
        new TextRun({
          text: block.authority,
          size: 28,
          font: { name: 'Times New Roman', eastAsia: '仿宋' },
        }),
      ],
      alignment: AlignmentType.RIGHT,
      indent: { right: 720 }, // ~1cm from right
      spacing: { before: 60, after: 0, line: 312 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: block.date,
          size: 28,
          font: { name: 'Times New Roman', eastAsia: '仿宋' },
        }),
      ],
      alignment: AlignmentType.RIGHT,
      indent: { right: 720 },
      spacing: { before: 0, after: 60, line: 312 },
    }),
  ];
}

function renderAttachmentNote(
  block: ContentBlock & { attachments: string[] },
): Paragraph[] {
  const prefix = '附件：';
  const text = prefix + block.attachments.map((a, i) => `${i + 1}. ${a}`).join('  ');
  return [
    new Paragraph({
      children: [
        new TextRun({
          text,
          size: 28,
          font: { name: 'Times New Roman', eastAsia: '仿宋' },
        }),
      ],
      alignment: AlignmentType.LEFT,
      spacing: { before: 60, after: 60, line: 312 },
    }),
  ];
}

// ── Block Dispatcher ──────────────────────────────────────

/**
 * The core block dispatcher — same pattern as your generate_docx.mjs.
 * Routes each ContentBlock to its renderer based on `type`.
 */
function processBlock(
  block: ContentBlock,
  template: DocumentTemplate,
): (Paragraph | Table)[] {
  switch (block.type) {
    case 'heading1':
    case 'heading2':
    case 'heading3':
    case 'heading4':
      return [renderHeading(template, block)];

    case 'paragraph':
    case 'paragraph_no_indent':
      return [renderParagraph(template, block)];

    case 'centered_text':
      return [renderCenteredText(template, block)];

    case 'equation':
    case 'equation_numbered':
      return [renderEquation(template, block)];

    case 'table':
      return renderTable(template, block);

    case 'figure':
      return [renderFigure(template, block)];

    case 'spacer':
      return renderSpacer(template, block.lines);

    case 'page_break':
      return [renderPageBreak()];

    case 'horizontal_rule':
      return [renderHorizontalRule()];

    case 'red_header':
      return [renderRedHeader(block)];

    case 'document_number':
      return [renderDocumentNumber(block)];

    case 'recipient_line':
      return [renderRecipientLine(block)];

    case 'signature_block':
      return renderSignatureBlock(block);

    case 'attachment_note':
      return renderAttachmentNote(block);

    case 'list_item':
    case 'code_block':
    case 'blockquote':
      // TODO: Implement in V2
      return [new Paragraph({
        children: [new TextRun({ text: `[${block.type}: ${(block as any).text}]`, size: 24 })],
      })];

    default:
      console.warn(`Unknown block type: ${(block as any).type}`);
      return [];
  }
}

// ── Main Render Function ──────────────────────────────────

/**
 * Render a ThesisDocument to a .docx file using the provided DocumentTemplate.
 * This is the main entry point.
 */
export async function renderDocument(options: RenderOptions): Promise<Buffer> {
  const { template, document: doc, headerText, showPageNumbers = true } = options;

  const sections: any[] = [];
  const children: (Paragraph | Table)[] = [];

  // ── Cover Page ─────────────────────────────────────────
  for (const block of doc.cover) {
    children.push(...processBlock(block, template));
  }

  // ── Body Sections ──────────────────────────────────────
  // Flatten sections recursively
  function flattenSections(sections: typeof doc.sections): ContentBlock[] {
    const result: ContentBlock[] = [];
    for (const section of sections) {
      // Section heading
      result.push({
        type: 'heading1',
        text: section.title,
        number: section.number,
        id: section.id,
      });
      // Section content
      result.push(...section.content);
      // Subsections
      if (section.subsections) {
        for (const sub of section.subsections) {
          result.push({
            type: 'heading2',
            text: sub.title,
            number: sub.number,
            id: sub.id,
          });
          result.push(...sub.content);
        }
      }
    }
    return result;
  }

  if (doc.sections && doc.sections.length > 0) {
    children.push(...processBlocks(flattenSections(doc.sections), template));
  }

  // ── Back Matter ────────────────────────────────────────
  if (doc.backMatter) {
    if (doc.backMatter.references && doc.backMatter.references.length > 0) {
      children.push(
        renderHeading(template, { type: 'heading1', text: '参考文献' }),
        ...doc.backMatter.references.map(ref =>
          renderParagraph(template, {
            type: 'paragraph_no_indent',
            text: `[${ref.id}] ${ref.text}`,
          }),
        ),
      );
    }
  }

  // Headers / Footers
  const headerStr = headerText || `${template.meta.organization}学位论文`;
  const headers = {
    default: new Header({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: headerStr,
              size: 21,
              font: { name: 'Times New Roman', eastAsia: '宋体' },
            }),
          ],
          alignment: AlignmentType.CENTER,
        }),
      ],
    }),
  };

  const footers = showPageNumbers ? {
    default: new Footer({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              children: [PageNumber.CURRENT],
              size: 21,
              font: 'Times New Roman',
            }),
          ],
          alignment: AlignmentType.CENTER,
        }),
      ],
    }),
  } : undefined;

  // ── Build Document ─────────────────────────────────────
  const wordDoc = new Document({
    sections: [{
      properties: {
        page: {
          size: {
            width: template.page.width,
            height: template.page.height,
          },
          margin: {
            top: template.page.margins.top,
            bottom: template.page.margins.bottom,
            left: template.page.margins.left,
            right: template.page.margins.right,
          },
        },
      },
      headers,
      footers,
      children,
    }],
  });

  const buffer = await Packer.toBuffer(wordDoc);

  if (options.outputPath) {
    writeFileSync(options.outputPath, buffer);
  }

  return buffer;
}

// ── Legacy API (backward compat with existing JSON format) ──

/**
 * Render using the legacy flat JSON format (backward compatible).
 * This wraps the old cover_blocks/body_blocks into a ThesisDocument.
 */
export async function renderLegacy(
  legacy: LegacyDocumentJSON,
  template: DocumentTemplate,
  outputPath: string,
): Promise<Buffer> {
  const doc: ThesisDocument = {
    type: 'thesis',
    meta: { title: legacy.title },
    cover: legacy.cover_blocks,
    sections: [
      {
        id: 'body',
        type: 'chapter',
        title: legacy.title,
        content: legacy.body_blocks,
      },
    ],
  };

  return renderDocument({
    outputPath,
    template,
    document: doc,
    headerText: legacy.report_header,
  });
}

// ── Utility ───────────────────────────────────────────────

function mapAlignment(align: string): (typeof AlignmentType)[keyof typeof AlignmentType] {
  switch (align) {
    case 'center': return AlignmentType.CENTER;
    case 'left': return AlignmentType.LEFT;
    case 'right': return AlignmentType.RIGHT;
    case 'justified': return AlignmentType.JUSTIFIED;
    case 'distribute': return AlignmentType.DISTRIBUTE;
    default: return AlignmentType.JUSTIFIED;
  }
}

/**
 * Process an array of blocks (for flattening section trees)
 */
function processBlocks(
  blocks: ContentBlock[],
  template: DocumentTemplate,
): (Paragraph | Table)[] {
  return blocks.flatMap(block => processBlock(block, template));
}

// ── Quick-start helpers ───────────────────────────────────

/**
 * Create a minimal USTB template (matching your existing format).
 * Useful as a default when no template file is provided.
 */
export function createUSTBTemplate(): DocumentTemplate {
  return {
    meta: {
      organization: '北京科技大学',
      name: '博士/硕士学位论文模板',
      documentType: 'thesis' as const,
      parserVersion: '0.1.0',
      parsedAt: new Date().toISOString(),
    },
    page: {
      width: convertMillimetersToTwip(210),
      height: convertMillimetersToTwip(297),
      margins: {
        top: convertMillimetersToTwip(30),
        bottom: convertMillimetersToTwip(20),
        left: convertMillimetersToTwip(30),
        right: convertMillimetersToTwip(30),
      },
    },
    styles: {
      ...DEFAULT_STYLES,
    },
    styleRoles: {
      'heading1': 'heading1',
      'heading2': 'heading2',
      'heading3': 'heading3',
      'Normal': 'paragraph',
      'CoverTitle': 'centered_text',
    },
  };
}
