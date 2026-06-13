// ============================================================
// @openthesis/template-parser — DOCX Template Style Extractor
// ============================================================
// Parses a .docx thesis template and extracts its formatting
// rules into a structured JSON DSL (DocumentTemplate).
// ============================================================

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import type {
  DocumentTemplate,
  TemplateMeta,
  PageSettings,
  ParagraphStyle,
  FontSettings,
  ParagraphFormatting,
  BlockType,
} from '@openthesis/document-schema';

// ── DOCX XML Namespaces ───────────────────────────────────
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
};

// Create XML parser with namespace-aware settings
function createParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name: string) =>
      ['w:p', 'w:r', 'w:t', 'w:tab', 'w:pPr', 'w:rPr', 'w:style',
       'w:tbl', 'w:tr', 'w:tc', 'w:sectPr'].includes(name),
    removeNSPrefix: false,
  });
}

// ── Main Entry Point ──────────────────────────────────────

/**
 * Parse a .docx file and extract the thesis template definition.
 * @param buffer — Raw .docx file content (as Buffer)
 * @param meta — Template metadata (university, degree, etc.)
 * @returns DocumentTemplate — structured style DSL
 */
export async function parseTemplate(
  buffer: Buffer | ArrayBuffer,
  metaOverrides?: Partial<TemplateMeta>,
): Promise<DocumentTemplate> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = createParser();

  // 1. Parse styles.xml
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  if (!stylesXml) throw new Error('No styles.xml found in template — is this a valid .docx?');

  const stylesParsed = parser.parse(stylesXml);
  const rawStyles = extractRawStyles(stylesParsed);

  // 2. Resolve style inheritance chain
  const resolvedStyles = resolveStyleInheritance(rawStyles);

  // 3. Parse document.xml for page settings
  const docXml = await zip.file('word/document.xml')?.async('string');
  const pageSettings = docXml
    ? extractPageSettings(parser.parse(docXml))
    : getDefaultPageSettings();

  // 4. Map styles to semantic roles (heuristic detection)
  const styleRoles = detectStyleRoles(resolvedStyles, Object.keys(rawStyles));

  // 5. Build template
  const meta: TemplateMeta = {
    organization: metaOverrides?.organization || 'Unknown Organization',
    name: metaOverrides?.name || 'Untitled Template',
    documentType: metaOverrides?.documentType || 'thesis',
    parserVersion: '0.1.0',
    parsedAt: new Date().toISOString(),
    ...metaOverrides,
  };

  // Convert to our style format
  const styles: Record<string, ParagraphStyle> = {};
  for (const [name, raw] of Object.entries(resolvedStyles)) {
    styles[name] = rawToParagraphStyle(raw);
  }

  return {
    meta,
    page: pageSettings,
    styles,
    styleRoles,
    styleInheritance: buildInheritanceMap(rawStyles),
  };
}

// ── Style Extraction ──────────────────────────────────────

interface RawStyle {
  styleId: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  basedOn?: string;
  /** Paragraph properties */
  pPr?: RawParagraphProps;
  /** Run (character) properties */
  rPr?: RawRunProps;
}

interface RawParagraphProps {
  alignment?: string;
  indent?: { firstLine?: number; left?: number; right?: number };
  spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  outlineLvl?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
}

interface RawRunProps {
  fontName?: string;
  eastAsiaFont?: string;
  fontSize?: number;        // in half-points
  fontSizeCs?: number;      // complex script font size
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

function extractRawStyles(parsed: any): Record<string, RawStyle> {
  const styles: Record<string, RawStyle> = {};
  const styleElements = parsed?.['w:styles']?.['w:style'] || [];

  for (const el of styleElements) {
    const styleId = el['@_w:styleId'];
    if (!styleId) continue;

    const type = el['@_w:type'] || 'paragraph';
    const basedOn = el['w:basedOn']?.['@_w:val'];

    const raw: RawStyle = {
      styleId,
      name: el['w:name']?.['@_w:val'] || styleId,
      type,
      basedOn,
      pPr: extractParagraphProps(el['w:pPr']),
      rPr: extractRunProps(el['w:rPr']),
    };

    styles[styleId] = raw;
  }

  // Also extract document defaults (DocDefaults)
  const docDefaults = parsed?.['w:styles']?.['w:docDefaults'];
  if (docDefaults) {
    const defaultRPr = extractRunProps(docDefaults['w:rPrDefault']?.['w:rPr']);
    const defaultPPr = extractParagraphProps(docDefaults['w:pPrDefault']?.['w:pPr']);
    styles['_defaults'] = {
      styleId: '_defaults',
      name: 'Document Defaults',
      type: 'paragraph',
      pPr: defaultPPr,
      rPr: defaultRPr,
    };
  }

  return styles;
}

function extractParagraphProps(pPr: any): RawParagraphProps | undefined {
  if (!pPr) return undefined;

  const props: RawParagraphProps = {};

  // Alignment
  const jc = pPr['w:jc']?.['@_w:val'];
  if (jc) props.alignment = jc;

  // Indentation
  const ind = pPr['w:ind'];
  if (ind) {
    props.indent = {};
    const firstLine = parseInt(ind['@_w:firstLine'] || ind['@_w:firstLineChars']);
    const left = parseInt(ind['@_w:left'] || ind['@_w:leftChars']);
    const right = parseInt(ind['@_w:right'] || ind['@_w:rightChars']);
    if (!isNaN(firstLine)) props.indent.firstLine = firstLine;
    if (!isNaN(left)) props.indent.left = left;
    if (!isNaN(right)) props.indent.right = right;
  }

  // Spacing
  const spacing = pPr['w:spacing'];
  if (spacing) {
    props.spacing = {};
    const before = parseInt(spacing['@_w:before']);
    const after = parseInt(spacing['@_w:after']);
    const line = parseInt(spacing['@_w:line']);
    if (!isNaN(before)) props.spacing.before = before;
    if (!isNaN(after)) props.spacing.after = after;
    if (!isNaN(line)) {
      props.spacing.line = line;
      props.spacing.lineRule = spacing['@_w:lineRule'] || 'auto';
    }
  }

  // Outline level
  const outlineLvl = pPr['w:outlineLvl']?.['@_w:val'];
  if (outlineLvl !== undefined) props.outlineLvl = parseInt(outlineLvl);

  // Keep with next / keep lines
  if (pPr['w:keepNext']) props.keepNext = true;
  if (pPr['w:keepLines']) props.keepLines = true;
  if (pPr['w:pageBreakBefore']) props.pageBreakBefore = true;

  return Object.keys(props).length > 0 ? props : undefined;
}

function extractRunProps(rPr: any): RawRunProps | undefined {
  if (!rPr) return undefined;

  const props: RawRunProps = {};

  // Font names
  const rFonts = rPr['w:rFonts'];
  if (rFonts) {
    if (rFonts['@_w:ascii']) props.fontName = rFonts['@_w:ascii'];
    if (rFonts['@_w:hAnsi']) props.fontName = props.fontName || rFonts['@_w:hAnsi'];
    if (rFonts['@_w:eastAsia']) props.eastAsiaFont = rFonts['@_w:eastAsia'];
  }

  // Font size (half-points)
  const sz = rPr['w:sz']?.['@_w:val'];
  if (sz) props.fontSize = parseInt(sz);
  const szCs = rPr['w:szCs']?.['@_w:val'];
  if (szCs) props.fontSizeCs = parseInt(szCs);

  // Bold
  if (rPr['w:b']) props.bold = rPr['w:b']['@_w:val'] !== '0' && rPr['w:b']['@_w:val'] !== 'false';
  // Italic
  if (rPr['w:i']) props.italic = rPr['w:i']['@_w:val'] !== '0' && rPr['w:i']['@_w:val'] !== 'false';
  // Underline
  if (rPr['w:u']) props.underline = rPr['w:u']['@_w:val'] !== 'none';
  // Color
  if (rPr['w:color']) props.color = rPr['w:color']['@_w:val'];

  return Object.keys(props).length > 0 ? props : undefined;
}

// ── Style Inheritance Resolution ──────────────────────────

/**
 * Resolve the style inheritance chain.
 * In DOCX, styles can be basedOn other styles and only define
 * overrides. We recursively merge all inherited properties.
 */
function resolveStyleInheritance(raw: Record<string, RawStyle>): Record<string, RawStyle> {
  const resolved: Record<string, RawStyle> = {};
  const defaults = raw['_defaults'];

  for (const [id, style] of Object.entries(raw)) {
    if (id === '_defaults') continue;
    resolved[id] = resolveSingleStyle(id, style, raw, new Set(), defaults);
  }

  return resolved;
}

function resolveSingleStyle(
  id: string,
  style: RawStyle,
  all: Record<string, RawStyle>,
  visited: Set<string>,
  defaults?: RawStyle,
): RawStyle {
  // Prevent infinite loops
  if (visited.has(id)) return style;
  visited.add(id);

  // No inheritance
  if (!style.basedOn || !all[style.basedOn]) {
    // Apply defaults at the bottom of the chain
    if (defaults) {
      return mergeStyles(defaults, style);
    }
    return style;
  }

  // Recursively resolve parent
  const parent = resolveSingleStyle(style.basedOn, all[style.basedOn], all, visited, defaults);
  return mergeStyles(parent, style);
}

function mergeStyles(base: RawStyle, override: RawStyle): RawStyle {
  return {
    ...base,
    ...override,
    styleId: override.styleId,  // keep the override's ID
    name: override.name || base.name,
    pPr: { ...base.pPr, ...override.pPr },
    rPr: { ...base.rPr, ...override.rPr },
  };
}

function buildInheritanceMap(raw: Record<string, RawStyle>): Record<string, string> {
  const map: Record<string, string> = {};
  for (const [id, style] of Object.entries(raw)) {
    if (style.basedOn) {
      map[id] = style.basedOn;
    }
  }
  return map;
}

// ── Style Conversion ──────────────────────────────────────

function rawToParagraphStyle(raw: RawStyle): ParagraphStyle {
  return {
    font: {
      name: raw.rPr?.fontName || 'Times New Roman',
      eastAsia: raw.rPr?.eastAsiaFont || '宋体',
      size: raw.rPr?.fontSize || 24,  // default 12pt if not specified
      bold: raw.rPr?.bold,
      italic: raw.rPr?.italic,
      underline: raw.rPr?.underline,
      color: raw.rPr?.color,
    },
    paragraph: {
      alignment: mapAlignment(raw.pPr?.alignment),
      firstLineIndent: raw.pPr?.indent?.firstLine ?? 0,
      leftIndent: raw.pPr?.indent?.left,
      rightIndent: raw.pPr?.indent?.right,
      spaceBefore: raw.pPr?.spacing?.before,
      spaceAfter: raw.pPr?.spacing?.after,
      outlineLevel: raw.pPr?.outlineLvl,
    },
    lineSpacing: raw.pPr?.spacing?.line,
  };
}

function mapAlignment(align?: string): ParagraphFormatting['alignment'] {
  switch (align) {
    case 'left': return 'left';
    case 'center': return 'center';
    case 'right': return 'right';
    case 'both': return 'justified';
    case 'distribute': return 'distribute';
    default: return 'justified';
  }
}

// ── Page Settings Extraction ───────────────────────────────

function extractPageSettings(parsed: any): PageSettings {
  const body = parsed?.['w:document']?.['w:body'];
  if (!body) return getDefaultPageSettings();

  // Look for section properties (sectPr) — usually last child of body
  const sectPr = body['w:sectPr']?.[0] || body['w:sectPr'];
  if (!sectPr) return getDefaultPageSettings();

  const pgSz = sectPr['w:pgSz'];
  const pgMar = sectPr['w:pgMar'];

  let width = 11906;   // A4 default in twips
  let height = 16838;
  if (pgSz) {
    const w = parseInt(pgSz['@_w:w']);
    const h = parseInt(pgSz['@_w:h']);
    if (!isNaN(w)) width = w;
    if (!isNaN(h)) height = h;
  }

  const margins = {
    top: 1440,     // 1 inch default
    bottom: 1440,
    left: 1800,    // 1.25 inch default
    right: 1800,
  };
  if (pgMar) {
    const t = parseInt(pgMar['@_w:top']);
    const b = parseInt(pgMar['@_w:bottom']);
    const l = parseInt(pgMar['@_w:left']);
    const r = parseInt(pgMar['@_w:right']);
    if (!isNaN(t)) margins.top = t;
    if (!isNaN(b)) margins.bottom = b;
    if (!isNaN(l)) margins.left = l;
    if (!isNaN(r)) margins.right = r;
  }

  const headerDistance = sectPr['w:headerReference']
    ? undefined  // we don't extract distance here, just mark existence
    : undefined;
  const footerDistance = sectPr['w:footerReference'] ? undefined : undefined;

  return { width, height, margins, headerDistance, footerDistance };
}

function getDefaultPageSettings(): PageSettings {
  return {
    width: 11906,   // A4
    height: 16838,
    margins: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
  };
}

// ── Semantic Role Detection ───────────────────────────────

/**
 * Heuristically detect the semantic role of each style.
 * This uses style name matching (Chinese + English) to guess
 * what each style is used for.
 *
 * In V4, this would be replaced by ML/layout-based detection.
 * For V1, we use heuristics based on common Chinese thesis patterns.
 */
function detectStyleRoles(
  styles: Record<string, RawStyle>,
  styleNames: string[],
): Record<string, BlockType> {
  const roles: Record<string, BlockType> = {};

  // Heuristic patterns for common Chinese thesis style naming
  const patterns: Array<{ regex: RegExp; role: BlockType }> = [
    // Chapter titles
    { regex: /^(标题\s*1|Heading\s*1|Chapter|第.*章|h1|标题1|章标题)$/i, role: 'heading1' },
    // Section titles
    { regex: /^(标题\s*2|Heading\s*2|Section|h2|标题2|节标题)$/i, role: 'heading2' },
    // Subsection titles
    { regex: /^(标题\s*3|Heading\s*3|Subsection|h3|标题3|小节标题)$/i, role: 'heading3' },
    { regex: /^(标题\s*4|Heading\s*4|h4|标题4)$/i, role: 'heading4' },
    // Body text
    { regex: /^(正文|Normal|Body|body\s*text|正文文本|普通)$/i, role: 'paragraph' },
    // Cover title
    { regex: /^(封面|Cover|Title|论文题目|题目)$/i, role: 'centered_text' },
    // Table caption
    { regex: /^(表|Table|题注)/i, role: 'heading3' },
  ];

  for (const name of styleNames) {
    const raw = styles[name];
    if (!raw) continue;

    let matched = false;
    for (const { regex, role } of patterns) {
      if (regex.test(name) || regex.test(raw.name)) {
        roles[name] = role;
        matched = true;
        break;
      }
    }

    // Fallback: use outline level
    if (!matched && raw.pPr?.outlineLvl !== undefined) {
      const lvl = raw.pPr.outlineLvl;
      if (lvl >= 0 && lvl <= 3) {
        roles[name] = `heading${lvl + 1}` as BlockType;
        matched = true;
      }
    }

    // Default: paragraph
    if (!matched) {
      roles[name] = 'paragraph';
    }
  }

  return roles;
}

// ── Utility Exports ───────────────────────────────────────

export { extractRawStyles, resolveStyleInheritance, detectStyleRoles };
