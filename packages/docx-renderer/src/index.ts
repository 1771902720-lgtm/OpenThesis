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
  ImageRun,
} from 'docx';
import type {
  DocumentTemplate,
  ThesisDocument,
  JournalArticle,
  JournalSection,
  OfficialDocument,
  ContentBlock,
  ParagraphStyle,
  BlockType,
  LegacyDocumentJSON,
  OpenThesisDocument,
} from '@openthesis/document-schema';
import { writeFileSync, readFileSync, existsSync } from 'fs';
import { resolve, extname, isAbsolute, dirname } from 'path';

// ── Renderer Config ───────────────────────────────────────

export interface RenderOptions {
  /** Path to write the output .docx */
  outputPath: string;
  /** Parsed template (from template-parser) */
  template: DocumentTemplate;
  /** Content to render */
  document: OpenThesisDocument;
  /** Header text (defaults to template.meta.organization + '学位论文') */
  headerText?: string;
  /** Show page numbers in footer */
  showPageNumbers?: boolean;
  /** Base directory of content file for relative paths (e.g. image paths) */
  contentDir?: string;
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
  const style = resolveStyle(_template, 'equation');
  const equationText = block.latex;
  const numberSuffix = block.number ? `    (${block.number})` : '';

  return new Paragraph({
    children: [
      new TextRun({
        text: equationText + numberSuffix,
        size: style.font.size || 24,
        font: {
          name: style.font.name || 'Times New Roman',
          eastAsia: style.font.eastAsia || '宋体',
        },
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
          bottom: BO,
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

function getImageSize(buffer: Buffer): { width: number; height: number } | null {
  // Check PNG signature
  if (buffer.readUInt32BE(0) === 0x89504E47 && buffer.readUInt32BE(4) === 0x0D0A1A0A) {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20)
    };
  }
  
  // Check JPEG signature
  if (buffer.readUInt16BE(0) === 0xFFD8) {
    let offset = 2;
    while (offset + 2 < buffer.length) {
      const marker = buffer.readUInt16BE(offset);
      offset += 2;
      
      if (marker >= 0xFFC0 && marker <= 0xFFCF && marker !== 0xFFC4 && marker !== 0xFFC8 && marker !== 0xFFCC) {
        // SOF segment: need at least 7 bytes (length(2) + precision(1) + height(2) + width(2))
        if (offset + 7 > buffer.length) break;
        return {
          height: buffer.readUInt16BE(offset + 3),
          width: buffer.readUInt16BE(offset + 5)
        };
      }
      
      if (offset + 2 > buffer.length) break;
      const segLength = buffer.readUInt16BE(offset);
      if (segLength < 2) break; // invalid segment length
      offset += segLength;
    }
  }
  
  // Check GIF signature
  if (buffer.readUInt32BE(0) === 0x47494638 && (buffer.readUInt16BE(4) === 0x3761 || buffer.readUInt16BE(4) === 0x3961)) {
    return {
      width: buffer.readUInt16LE(6),
      height: buffer.readUInt16LE(8)
    };
  }

  return null;
}

function renderFigure(
  _template: DocumentTemplate,
  block: ContentBlock & { caption: string; path: string; width?: number; height?: number },
  contentDir?: string,
): Paragraph {
  const ext = extname(block.path).toLowerCase();
  let type: 'png' | 'jpg' | 'gif' | 'bmp' = 'png';
  if (ext === '.jpg' || ext === '.jpeg') type = 'jpg';
  else if (ext === '.gif') type = 'gif';
  else if (ext === '.bmp') type = 'bmp';

  let imgPath = block.path;
  if (contentDir && !isAbsolute(imgPath)) {
    imgPath = resolve(contentDir, imgPath);
  }

  if (existsSync(imgPath)) {
    try {
      const buffer = readFileSync(imgPath);
      const size = getImageSize(buffer);
      
      let width = block.width;
      let height = block.height;
      
      if (size) {
        if (!width && !height) {
          // Default to a max width of 450px while keeping aspect ratio
          const targetWidth = Math.min(size.width, 450);
          const ratio = size.width / size.height;
          width = targetWidth;
          height = targetWidth / ratio;
        } else if (width && !height) {
          const ratio = size.width / size.height;
          height = width / ratio;
        } else if (!width && height) {
          const ratio = size.width / size.height;
          width = height * ratio;
        }
      } else {
        width = width || 400;
        height = height || 300;
      }

      return new Paragraph({
        children: [
          new ImageRun({
            data: buffer,
            transformation: {
              width: width!,
              height: height!,
            },
            type,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120, line: 312 },
      });
    } catch (e: any) {
      console.warn(`Failed to embed image at ${imgPath}: ${e.message}`);
    }
  }

  return new Paragraph({
    children: [
      new TextRun({
        text: `[图: ${block.caption} (未找到图片: ${block.path})]`,
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
  const prefix = block.recipientType === 'cc' ? '抄送：' : '';
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
  contentDir?: string,
): (Paragraph | Table)[] {
  switch (block.type) {
    case 'heading1':
    case 'heading2':
    case 'heading3':
    case 'heading4':
      return [renderHeading(template, block as any)];

    case 'paragraph':
    case 'paragraph_no_indent':
      return [renderParagraph(template, block as any)];

    case 'centered_text':
      return [renderCenteredText(template, block as any)];

    case 'equation':
    case 'equation_numbered':
      return [renderEquation(template, block as any)];

    case 'table':
      return renderTable(template, block as any);

    case 'figure':
      return [renderFigure(template, block as any, contentDir)];

    case 'spacer':
      return renderSpacer(template, block.lines);

    case 'page_break':
      return [renderPageBreak()];

    case 'horizontal_rule':
      return [renderHorizontalRule()];

    case 'red_header':
      return [renderRedHeader(block as any)];

    case 'document_number':
      return [renderDocumentNumber(block as any)];

    case 'recipient_line':
      return [renderRecipientLine(block as any)];

    case 'signature_block':
      return renderSignatureBlock(block as any);

    case 'attachment_note':
      return renderAttachmentNote(block as any);

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

// ── Structured Renderers ──────────────────────────────────

function renderThesisDocument(
  doc: ThesisDocument,
  template: DocumentTemplate,
  contentDir?: string,
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];

  // Cover Page
  for (const block of doc.cover) {
    children.push(...processBlock(block, template, contentDir));
  }

  const HEADING_TYPES = ['heading1', 'heading2', 'heading3', 'heading4'] as const;

  // Body Sections — recursively flatten to support arbitrary nesting depth
  function flattenSections(sections: typeof doc.sections, depth: number = 0): ContentBlock[] {
    const result: ContentBlock[] = [];
    const headingType = HEADING_TYPES[Math.min(depth, HEADING_TYPES.length - 1)];
    for (const section of sections) {
      result.push({
        type: headingType,
        text: section.title,
        number: section.number,
        id: section.id,
      } as ContentBlock);
      result.push(...section.content);
      if (section.subsections && section.subsections.length > 0) {
        result.push(...flattenSections(section.subsections, depth + 1));
      }
    }
    return result;
  }

  if (doc.sections && doc.sections.length > 0) {
    children.push(...processBlocks(flattenSections(doc.sections), template, contentDir));
  }

  // Back Matter
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

  return children;
}

function renderJournalArticle(
  doc: JournalArticle,
  template: DocumentTemplate,
  contentDir?: string,
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];

  // Title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: doc.meta.title,
          bold: true,
          size: 32,
          font: { name: 'Times New Roman', eastAsia: '黑体' },
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120, line: 312 },
    })
  );

  // Authors
  if (doc.meta.authors && doc.meta.authors.length > 0) {
    const authorRuns: TextRun[] = [];
    doc.meta.authors.forEach((author, index) => {
      authorRuns.push(
        new TextRun({
          text: author.name + (author.isCorresponding ? '*' : ''),
          size: 21,
          font: { name: 'Times New Roman', eastAsia: '宋体' },
        })
      );
      
      if (index < doc.meta.authors.length - 1) {
        authorRuns.push(new TextRun({ text: '  ', size: 21 }));
      }
    });

    children.push(
      new Paragraph({
        children: authorRuns,
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60, line: 312 },
      })
    );
  }

  // Affiliations
  const allAffiliations: string[] = [];
  doc.meta.authors.forEach(author => {
    author.affiliations.forEach(aff => {
      const affStr = `${aff.institution}${aff.department ? ', ' + aff.department : ''}${aff.city ? ', ' + aff.city : ''}`;
      if (!allAffiliations.includes(affStr)) {
        allAffiliations.push(affStr);
      }
    });
  });

  if (allAffiliations.length > 0) {
    const affText = allAffiliations.map((aff, i) => `(${i + 1}) ${aff}`).join('  ');
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: affText,
            size: 18,
            italics: true,
            font: { name: 'Times New Roman', eastAsia: '宋体' },
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120, line: 312 },
      })
    );
  }

  // Abstract Block
  if (doc.meta.abstract) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: '摘  要：', bold: true, size: 20, font: { name: 'Times New Roman', eastAsia: '黑体' } }),
          new TextRun({ text: doc.meta.abstract, size: 20, font: { name: 'Times New Roman', eastAsia: '宋体' } }),
        ],
        indent: { left: 720, right: 720 },
        spacing: { before: 120, after: 60, line: 240 },
        alignment: AlignmentType.JUSTIFIED,
      })
    );
  }

  // Keywords
  if (doc.meta.keywords && doc.meta.keywords.length > 0) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: '关键词：', bold: true, size: 20, font: { name: 'Times New Roman', eastAsia: '黑体' } }),
          new TextRun({ text: doc.meta.keywords.join('；'), size: 20, font: { name: 'Times New Roman', eastAsia: '宋体' } }),
        ],
        indent: { left: 720, right: 720 },
        spacing: { before: 60, after: 240, line: 240 },
        alignment: AlignmentType.LEFT,
      })
    );
  }

  // Divider line or spacer before body
  children.push(new Paragraph({ children: [], spacing: { after: 120 } }));

  const HEADING_TYPES = ['heading1', 'heading2', 'heading3', 'heading4'] as const;

  // Sections — recursively flatten to support arbitrary nesting depth
  function flattenJournalSections(sections: JournalSection[], depth: number = 0): ContentBlock[] {
    const result: ContentBlock[] = [];
    const headingType = HEADING_TYPES[Math.min(depth, HEADING_TYPES.length - 1)];
    for (const section of sections) {
      result.push({
        type: headingType,
        text: section.title,
        number: section.number,
        id: section.id,
      } as ContentBlock);
      result.push(...section.content);
      if (section.subsections && section.subsections.length > 0) {
        result.push(...flattenJournalSections(section.subsections, depth + 1));
      }
    }
    return result;
  }

  children.push(...processBlocks(flattenJournalSections(doc.sections), template, contentDir));

  // Back Matter
  if (doc.backMatter) {
    if (doc.backMatter.references && doc.backMatter.references.length > 0) {
      children.push(
        renderHeading(template, { type: 'heading1', text: 'References' }),
        ...doc.backMatter.references.map(ref =>
          renderParagraph(template, {
            type: 'paragraph_no_indent',
            text: `[${ref.id}] ${ref.text}`,
          })
        )
      );
    }
    
    if (doc.backMatter.acknowledgments) {
      children.push(
        renderHeading(template, { type: 'heading1', text: 'Acknowledgments' }),
        renderParagraph(template, { type: 'paragraph', text: doc.backMatter.acknowledgments })
      );
    }
  }

  return children;
}

function renderOfficialDocument(
  doc: OfficialDocument,
  template: DocumentTemplate,
  contentDir?: string,
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];

  // Urgency
  if (doc.meta.urgency && doc.meta.urgency !== 'normal') {
    const urgencyText = doc.meta.urgency === 'urgent' ? '急件' : '特急';
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: urgencyText,
            size: 24,
            bold: true,
            color: 'FF0000',
            font: { name: 'Times New Roman', eastAsia: '黑体' },
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { before: 120, after: 120 },
      })
    );
  }

  // Red Header
  children.push(
    renderRedHeader({
      type: 'red_header',
      text: doc.meta.issuingAuthority + '文件',
    })
  );

  // Document Number
  children.push(
    renderDocumentNumber({
      type: 'document_number',
      text: doc.meta.documentNumber,
    })
  );

  // Red horizontal line
  children.push(
    new Paragraph({
      children: [],
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: 'FF0000' } },
      spacing: { after: 240 },
    })
  );

  // Title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: doc.meta.title,
          size: 36,
          bold: true,
          font: { name: 'Times New Roman', eastAsia: '方正小标宋简体' },
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 240, line: 360 },
    })
  );

  // Recipients
  if (doc.meta.primaryRecipients && doc.meta.primaryRecipients.length > 0) {
    children.push(
      renderRecipientLine({
        type: 'recipient_line',
        text: doc.meta.primaryRecipients.join('、') + '：',
        recipientType: 'primary',
      })
    );
  }

  // Body
  children.push(...processBlocks(doc.body, template, contentDir));

  // Signature Block
  children.push(
    ...renderSignatureBlock({
      type: 'signature_block',
      authority: doc.meta.signatureAuthority || doc.meta.issuingAuthority,
      date: doc.meta.date,
    })
  );

  // Attachments
  if (doc.attachments && doc.attachments.length > 0) {
    children.push(
      ...renderAttachmentNote({
        type: 'attachment_note',
        attachments: doc.attachments.map(att => att.title),
      })
    );
    
    for (const att of doc.attachments) {
      if (att.content && att.content.length > 0) {
        children.push(renderPageBreak());
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `附件${att.order}：${att.title}`,
                size: 28,
                bold: true,
                font: { name: 'Times New Roman', eastAsia: '黑体' },
              }),
            ],
            spacing: { before: 240, after: 240, line: 312 },
          })
        );
        children.push(...processBlocks(att.content, template, contentDir));
      }
    }
  }

  // CC Recipients
  if (doc.meta.ccRecipients && doc.meta.ccRecipients.length > 0) {
    children.push(
      new Paragraph({
        children: [],
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
        spacing: { before: 240 },
      })
    );
    children.push(
      renderRecipientLine({
        type: 'recipient_line',
        text: '抄送：' + doc.meta.ccRecipients.join('、') + '。',
        recipientType: 'cc',
      })
    );
  }

  // Publishing Info
  const issuingDept = doc.meta.issuingDepartment || doc.meta.issuingAuthority;
  const issueDate = doc.meta.issueDate || doc.meta.date;
  children.push(
    new Paragraph({
      children: [],
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
      spacing: { before: 60 },
    })
  );
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${issuingDept}办公厅`,
          size: 18,
          font: { name: 'Times New Roman', eastAsia: '仿宋' },
        }),
        new TextRun({
          text: `\t\t\t\t\t\t${issueDate}印发`,
          size: 18,
          font: { name: 'Times New Roman', eastAsia: '仿宋' },
        }),
      ],
      spacing: { before: 60, after: 120 },
    })
  );

  return children;
}

// ── Main Render Function ──────────────────────────────────

/**
 * Render any structured OpenThesisDocument to a .docx file using the provided DocumentTemplate.
 * This is the main entry point.
 */
export async function renderDocument(options: RenderOptions): Promise<Buffer> {
  const { template, document: doc, headerText, showPageNumbers = true, contentDir } = options;

  let children: (Paragraph | Table)[] = [];

  if (doc.type === 'thesis') {
    children = renderThesisDocument(doc as ThesisDocument, template, contentDir);
  } else if (doc.type === 'journal') {
    children = renderJournalArticle(doc as JournalArticle, template, contentDir);
  } else if (doc.type === 'official') {
    children = renderOfficialDocument(doc as OfficialDocument, template, contentDir);
  } else {
    throw new Error(`Unsupported document type: "${(doc as any).type}". Expected 'thesis', 'journal', or 'official'.`);
  }

  // Headers / Footers
  const headerStr = headerText || (doc.type === 'official' ? '' : `${template.meta.organization}${doc.type === 'journal' ? '学术论文' : '学位论文'}`);
  const headers = doc.type === 'official' ? undefined : {
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
  const sectionProperties: any = {
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
  };

  // Section level columns
  if (template.page.columns && template.page.columns > 1) {
    sectionProperties.columns = {
      count: template.page.columns,
      space: template.page.columnGutter ?? 708,
    };
  }

  const wordDoc = new Document({
    sections: [{
      properties: sectionProperties,
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

  const contentDir = dirname(outputPath);

  return renderDocument({
    outputPath,
    template,
    document: doc,
    headerText: legacy.report_header,
    contentDir,
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
  contentDir?: string,
): (Paragraph | Table)[] {
  return blocks.flatMap(block => processBlock(block, template, contentDir));
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
