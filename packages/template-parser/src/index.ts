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
  ParagraphFormatting,
  LineSpacingRule,
  BlockType,
} from '@openthesis/document-schema';

const WORDPROCESSINGML_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * Rewrite the document's own WordprocessingML prefix to `w:`.
 *
 * Nothing in OOXML requires the prefix to be `w` — it is declared per document.
 * Every lookup in this file is written against the literal `w:`, so a
 * styles.xml binding the namespace to another prefix parsed to zero styles and
 * the template came out empty. The binding is declared on the root element, so
 * it is read from there and normalised before parsing.
 */
function normalizeWordPrefix(xml: string): string {
  const declaration = /xmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g;
  let prefix: string | undefined;
  for (let match = declaration.exec(xml); match; match = declaration.exec(xml)) {
    if (match[2] === WORDPROCESSINGML_NAMESPACE) { prefix = match[1]; break; }
  }
  if (!prefix || prefix === 'w') return xml;

  const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return xml
    .replace(new RegExp(`<${escaped}:`, 'g'), '<w:')
    .replace(new RegExp(`</${escaped}:`, 'g'), '</w:')
    .replace(new RegExp(`(\\s)${escaped}:`, 'g'), '$1w:');
}

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

  const stylesParsed = parser.parse(normalizeWordPrefix(stylesXml));
  const rawStyles = extractRawStyles(stylesParsed);

  // 2. Resolve style inheritance chain. A `basedOn` loop is malformed; break it
  // first so the merged result does not depend on where the walk started.
  const { styles: acyclicStyles, cycles } = breakStyleCycles(rawStyles);
  const resolvedStyles = resolveStyleInheritance(acyclicStyles);

  // 3. Parse document.xml for page settings
  const docXml = await zip.file('word/document.xml')?.async('string');
  const pageSections = docXml ? extractPageSections(parser.parse(normalizeWordPrefix(docXml))) : [];
  const pageSettings = pageSections.length > 0
    ? pageSections[pageSections.length - 1]
    : getDefaultPageSettings();

  // 4. Map styles to semantic roles (heuristic detection)
  const styleRoles = detectStyleRoles(resolvedStyles, Object.keys(rawStyles));

  // 5. Build template
  const meta: TemplateMeta = {
    ...metaOverrides,
    organization: metaOverrides?.organization?.trim() || 'Unknown Organization',
    name: metaOverrides?.name?.trim() || 'Untitled Template',
    documentType: metaOverrides?.documentType || 'thesis',
    parserVersion: '0.2.0',
    parsedAt: new Date().toISOString(),
  };

  // Convert to our style format
  const styles: Record<string, ParagraphStyle> = {};
  for (const [name, raw] of Object.entries(resolvedStyles)) {
    styles[name] = rawToParagraphStyle(raw);
  }

  const warnings = buildWarnings(rawStyles, resolvedStyles, pageSections, cycles);

  return {
    meta,
    page: pageSettings,
    ...(pageSections.length > 0 ? { pageSections } : {}),
    styles,
    styleRoles,
    styleInheritance: buildInheritanceMap(rawStyles),
    ...(warnings.length > 0 ? { warnings } : {}),
  };
}

/**
 * Non-fatal problems worth telling the caller about.
 *
 * These exist because a template that carries no formatting used to parse
 * "successfully" into a DocumentTemplate whose every style was identical, and
 * the first symptom was a document with headings that looked like body text.
 */
function buildWarnings(
  rawStyles: Record<string, RawStyle>,
  resolvedStyles: Record<string, RawStyle>,
  pageSections: PageSettings[],
  cycles: string[] = [],
): string[] {
  const warnings: string[] = [];
  const styleIds = Object.keys(resolvedStyles);

  if (styleIds.length === 0) {
    warnings.push('styles.xml defines no paragraph styles; every block will use the built-in defaults.');
    return warnings;
  }

  const missingParents = new Set<string>();
  for (const style of Object.values(rawStyles)) {
    if (style.basedOn && !rawStyles[style.basedOn]) {
      missingParents.add(`${style.styleId} → ${style.basedOn}`);
    }
  }
  if (missingParents.size > 0) {
    const shown = [...missingParents].slice(0, 5).join(', ');
    const more = missingParents.size > 5 ? ` (+${missingParents.size - 5} more)` : '';
    warnings.push(`styles inherit from styles this template does not define: ${shown}${more}.`);
  }

  if (cycles.length > 0) {
    const shown = cycles.slice(0, 3).join('; ');
    const more = cycles.length > 3 ? ` (+${cycles.length - 3} more)` : '';
    warnings.push(
      `styles inherit from each other in a loop (${shown}${more}); inheritance stops at the first style `
      + 'in each loop, so the formatting may not match the source document.',
    );
  }

  const declaresFormatting = styleIds.some(id => {
    const style = resolvedStyles[id];
    return style.rPr !== undefined || style.pPr !== undefined;
  });
  if (!declaresFormatting) {
    warnings.push(
      'No style declares any font or paragraph formatting, so this template carries no typography; '
      + 'every block will fall back to the built-in defaults. Re-check the source .docx, or author the '
      + 'styles explicitly.',
    );
  }

  if (pageSections.length > 1) {
    warnings.push(
      `The document declares ${pageSections.length} sections; only the last is used for page geometry, `
      + 'so cover or appendix page setups are not reproduced.',
    );
  }

  return warnings;
}

// ── Style Extraction ──────────────────────────────────────

interface RawStyle {
  styleId: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  /** True for the `w:default="1"` style of its type (usually Normal/正文). */
  default?: boolean;
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
    const pPrEl = firstElement(el['w:pPr']);

    // Chinese Word/WPS templates routinely put run properties (黑体, 三号, …)
    // inside `w:pPr/w:rPr` instead of a style-level `w:rPr`. Reading only the
    // direct child silently dropped them, which made every style look empty.
    // A style-level `w:rPr` wins when both are present.
    const nestedRPr = extractRunProps(pPrEl?.['w:rPr']);
    const directRPr = extractRunProps(el['w:rPr']);
    const rPr = mergeRunProps(nestedRPr, directRPr);

    const raw: RawStyle = {
      styleId,
      name: el['w:name']?.['@_w:val'] || styleId,
      type,
      default: el['@_w:default'] === '1' || el['@_w:default'] === 'true',
      basedOn,
      pPr: extractParagraphProps(el['w:pPr'], rPr?.fontSize),
      rPr,
    };

    styles[styleId] = raw;
  }

  // Also extract document defaults (DocDefaults)
  const docDefaults = parsed?.['w:styles']?.['w:docDefaults'];
  if (docDefaults) {
    const defaultRPr = extractRunProps(docDefaults['w:rPrDefault']?.['w:rPr']);
    const defaultPPr = extractParagraphProps(docDefaults['w:pPrDefault']?.['w:pPr'], defaultRPr?.fontSize);
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

/** Merge two run-property bags; `override` wins. Keys are only present when set. */
function mergeRunProps(base?: RawRunProps, override?: RawRunProps): RawRunProps | undefined {
  if (!base && !override) return undefined;
  return { ...base, ...override };
}

function extractParagraphProps(pPr: any, runSize?: number): RawParagraphProps | undefined {
  pPr = firstElement(pPr);
  if (!pPr) return undefined;

  const props: RawParagraphProps = {};

  // Alignment
  const jc = pPr['w:jc']?.['@_w:val'];
  if (jc) props.alignment = jc;

  // Indentation. `w:firstLineChars` / `w:leftChars` / `w:rightChars` are
  // hundredths of a *character*, not twips, and Chinese templates emit them
  // constantly. Converting needs the run size: twips = chars × halfPoints / 10.
  const ind = pPr['w:ind'];
  if (ind) {
    props.indent = {};
    const effectiveSize = runSize && runSize > 0 ? runSize : 24; // 12pt fallback
    const charsToTwips = (chars: number) => Math.round((chars * effectiveSize) / 10);

    const firstLine = parseInt(ind['@_w:firstLine']);
    const firstLineChars = parseInt(ind['@_w:firstLineChars']);
    const left = parseInt(ind['@_w:left']);
    const leftChars = parseInt(ind['@_w:leftChars']);
    const right = parseInt(ind['@_w:right']);
    const rightChars = parseInt(ind['@_w:rightChars']);

    if (!isNaN(firstLine)) props.indent.firstLine = firstLine;
    else if (!isNaN(firstLineChars)) props.indent.firstLine = charsToTwips(firstLineChars);

    if (!isNaN(left)) props.indent.left = left;
    else if (!isNaN(leftChars)) props.indent.left = charsToTwips(leftChars);

    if (!isNaN(right)) props.indent.right = right;
    else if (!isNaN(rightChars)) props.indent.right = charsToTwips(rightChars);
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
  rPr = firstElement(rPr);
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
  if (rPr['w:b'] !== undefined) {
    const value = rPr['w:b']?.['@_w:val'];
    props.bold = value !== '0' && value !== 'false';
  }
  // Italic
  if (rPr['w:i'] !== undefined) {
    const value = rPr['w:i']?.['@_w:val'];
    props.italic = value !== '0' && value !== 'false';
  }
  // Underline
  if (rPr['w:u'] !== undefined) {
    const value = rPr['w:u']?.['@_w:val'];
    props.underline = value !== 'none' && value !== '0' && value !== 'false';
  }
  // Color
  if (rPr['w:color']) props.color = rPr['w:color']['@_w:val'];

  return Object.keys(props).length > 0 ? props : undefined;
}

function firstElement<T>(value: T | T[] | undefined): T | undefined {
  return Array.isArray(value) ? value[0] : value;
}

// ── Style Inheritance Resolution ──────────────────────────

/**
 * Cut every `basedOn` loop at a canonical point.
 *
 * A style whose chain returns to itself is malformed, and the merge used to
 * depend on which of its styles happened to be resolved first. Dropping the
 * link that closes the loop — always at the earliest-declared member — makes
 * the result the same whichever entry point the walk uses, and the loop is
 * reported through the template's warnings.
 */
function breakStyleCycles(
  styles: Record<string, RawStyle>,
): { styles: Record<string, RawStyle>; cycles: string[] } {
  const ids = Object.keys(styles);
  const cut = new Set<string>();
  const cycles: string[] = [];

  for (const start of ids) {
    const path: string[] = [];
    let current: string | undefined = start;
    while (current && styles[current] && !path.includes(current) && !cut.has(current)) {
      path.push(current);
      current = styles[current].basedOn;
    }
    if (!current || !path.includes(current)) continue;

    const loop = path.slice(path.indexOf(current));
    cut.add(loop.reduce((a, b) => (ids.indexOf(a) <= ids.indexOf(b) ? a : b)));
    cycles.push(`${loop.join(' → ')} → ${loop[0]}`);
  }

  if (cut.size === 0) return { styles, cycles };

  const broken: Record<string, RawStyle> = {};
  for (const [id, style] of Object.entries(styles)) {
    broken[id] = cut.has(id) ? { ...style, basedOn: undefined } : style;
  }
  return { styles: broken, cycles };
}

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
    pPr: mergeParagraphProps(base.pPr, override.pPr),
    rPr: { ...base.rPr, ...override.rPr },
  };
}

/** Deep-merge paragraph properties so nested indent/spacing are not lost */
function mergeParagraphProps(
  base?: RawParagraphProps,
  override?: RawParagraphProps,
): RawParagraphProps | undefined {
  if (!base && !override) return undefined;
  if (!base) return override;
  if (!override) return base;
  return {
    ...base,
    ...override,
    indent: (base.indent || override.indent)
      ? { ...base.indent, ...override.indent }
      : undefined,
    spacing: (base.spacing || override.spacing)
      ? { ...base.spacing, ...override.spacing }
      : undefined,
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

/** Drop keys whose value is `undefined` so the JSON stays honest and compact. */
function definedOnly<T extends object>(value: T): T {
  return Object.fromEntries(
    Object.entries(value).filter(([, v]) => v !== undefined),
  ) as T;
}

/**
 * Convert resolved OOXML style properties into the public ParagraphStyle shape.
 *
 * Only properties the template actually declares are emitted. Inventing
 * fallbacks here (the previous `|| 24` / `|| '宋体'` / `?? 0`) made a silent
 * template indistinguishable from one that explicitly asked for 12pt 宋体,
 * which is why every style in a formatting-free template looked identical and
 * headings ended up rendered as body text.
 */
function rawToParagraphStyle(raw: RawStyle): ParagraphStyle {
  return definedOnly({
    font: definedOnly({
      name: raw.rPr?.fontName,
      eastAsia: raw.rPr?.eastAsiaFont,
      size: raw.rPr?.fontSize,
      bold: raw.rPr?.bold,
      italic: raw.rPr?.italic,
      underline: raw.rPr?.underline,
      color: raw.rPr?.color,
    }),
    paragraph: definedOnly({
      alignment: mapAlignment(raw.pPr?.alignment),
      firstLineIndent: raw.pPr?.indent?.firstLine,
      leftIndent: raw.pPr?.indent?.left,
      rightIndent: raw.pPr?.indent?.right,
      spaceBefore: raw.pPr?.spacing?.before,
      spaceAfter: raw.pPr?.spacing?.after,
      outlineLevel: raw.pPr?.outlineLvl,
    }),
    lineSpacing: raw.pPr?.spacing?.line,
    lineSpacingRule: toLineSpacingRule(raw.pPr?.spacing?.lineRule),
  });
}

/**
 * OOXML's `w:lineRule`: `auto` measures the line in 240ths, `exact` and
 * `atLeast` in twips. A value outside that vocabulary is dropped rather than
 * guessed at.
 */
function toLineSpacingRule(value?: string): LineSpacingRule | undefined {
  return value === 'auto' || value === 'exact' || value === 'atLeast' ? value : undefined;
}

function mapAlignment(align?: string): ParagraphFormatting['alignment'] | undefined {
  switch (align) {
    case 'left': return 'left';
    case 'center': return 'center';
    case 'right': return 'right';
    case 'both': return 'justified';
    case 'distribute': return 'distribute';
    default: return undefined;
  }
}

// ── Page Settings Extraction ───────────────────────────────

/**
 * Every section geometry in the document, in document order.
 *
 * In OOXML every section but the last keeps its properties in
 * `w:p/w:pPr/w:sectPr`; the final section's properties are a direct child of
 * `w:body`. Reading only the body-level element meant a template with a
 * distinct cover or a landscape appendix reported a single geometry.
 */
function extractPageSections(parsed: any): PageSettings[] {
  const body = parsed?.['w:document']?.['w:body'];
  if (!body) return [];

  const sections: PageSettings[] = [];

  const paragraphs = body['w:p'];
  for (const paragraph of Array.isArray(paragraphs) ? paragraphs : paragraphs ? [paragraphs] : []) {
    const pPr = firstElement(paragraph?.['w:pPr']);
    const sectPr = firstElement(pPr?.['w:sectPr']);
    if (sectPr) sections.push(sectPrToPageSettings(sectPr));
  }

  const bodySectPr = firstElement(body['w:sectPr']);
  if (bodySectPr) sections.push(sectPrToPageSettings(bodySectPr));

  return sections;
}

function sectPrToPageSettings(sectPr: any): PageSettings {
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

  // Extract header/footer distances from pgMar attributes
  let headerDistance: number | undefined;
  let footerDistance: number | undefined;
  if (pgMar) {
    const hd = parseInt(pgMar['@_w:header']);
    if (!isNaN(hd)) headerDistance = hd;
    const fd = parseInt(pgMar['@_w:footer']);
    if (!isNaN(fd)) footerDistance = fd;
  }

  const cols = sectPr['w:cols'];
  let columns = 1;
  let columnGutter: number | undefined;
  if (cols) {
    const num = parseInt(cols['@_w:num']);
    if (!isNaN(num)) columns = num;
    const space = parseInt(cols['@_w:space']);
    if (!isNaN(space)) columnGutter = space;
  }

  return { width, height, margins, headerDistance, footerDistance, columns, columnGutter };
}

function getDefaultPageSettings(): PageSettings {
  return {
    width: 11906,   // A4
    height: 16838,
    margins: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
    columns: 1,
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

  // Heuristic patterns for common Chinese thesis style naming.
  // Patterns are anchored on purpose: the previous unanchored `/^(表|Table|题注)/`
  // also matched Word's built-in `TableGrid` ("Table Grid") and
  // "Table of Contents", turning table and TOC styles into level-3 headings.
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
    // Equation styles — MathType emits "MT Converted Equation", WPS emits "公式".
    // Without this, no template style ever drove `equation` and every formula
    // silently fell back to the built-in defaults.
    { regex: /^(公式|Equation|MT\s*Converted\s*Equation|Math\s*Equation)$/i, role: 'equation' },
  ];

  // Unmatched paragraph styles still act as body text, but they are appended
  // after every explicit match so that the renderer's first-match lookup picks
  // a style that actually declares a body-text role.
  const defaultParagraph: string[] = [];
  const otherParagraph: string[] = [];

  for (const name of styleNames) {
    const raw = styles[name];
    if (!raw) continue;

    // Character, table and numbering styles are not paragraph-level blocks.
    // Assigning them a block role made `resolveStyle` pick whichever style
    // happened to appear first in the XML — e.g. `uChar` (a character style)
    // or `TOC3` supplying the body-text formatting.
    if (raw.type !== 'paragraph') continue;

    let matched = false;
    for (const { regex, role } of patterns) {
      if (regex.test(name) || regex.test(raw.name)) {
        roles[name] = role;
        matched = true;
        break;
      }
    }
    if (matched) continue;

    // Fallback: use outline level
    if (raw.pPr?.outlineLvl !== undefined) {
      const lvl = raw.pPr.outlineLvl;
      if (lvl >= 0 && lvl <= 3) {
        roles[name] = `heading${lvl + 1}` as BlockType;
        continue;
      }
    }

    (raw.default ? defaultParagraph : otherParagraph).push(name);
  }

  for (const name of defaultParagraph) roles[name] = 'paragraph';
  for (const name of otherParagraph) roles[name] = 'paragraph';

  return roles;
}

// ── Utility Exports ───────────────────────────────────────

export { extractRawStyles, resolveStyleInheritance, detectStyleRoles };
