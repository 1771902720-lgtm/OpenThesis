import type {
  ContentBlock,
  DocumentType,
  JournalArticle,
  JournalAuthor,
  JournalSection,
  OfficialDocCategory,
  OfficialDocument,
  OpenThesisDocument,
  ThesisDocument,
  ThesisSection,
} from '@openthesis/document-schema';

type FrontMatterValue = string | string[] | number | boolean;
export type MarkdownFrontMatter = Record<string, FrontMatterValue>;

export interface MarkdownParseOptions {
  documentType?: DocumentType;
  sourceName?: string;
  metadata?: MarkdownFrontMatter;
}

export interface MarkdownParseResult {
  document: OpenThesisDocument;
  frontMatter: MarkdownFrontMatter;
}

interface HeadingItem {
  kind: 'heading';
  level: number;
  text: string;
}

interface BlockItem {
  kind: 'block';
  block: ContentBlock;
}

type MarkdownItem = HeadingItem | BlockItem;

interface SectionNode {
  level: number;
  title: string;
  content: ContentBlock[];
  children: SectionNode[];
}

const DOCUMENT_TYPES: DocumentType[] = ['thesis', 'journal', 'official'];
const OFFICIAL_CATEGORIES: OfficialDocCategory[] = [
  '通知', '通告', '报告', '请示', '批复', '函', '纪要', '决定',
  '意见', '通报', '议案', '公告', '命令', '决议', '公报',
];

export function parseMarkdown(markdown: string, options: MarkdownParseOptions = {}): OpenThesisDocument {
  return parseMarkdownWithMetadata(markdown, options).document;
}

export function parseMarkdownWithMetadata(
  markdown: string,
  options: MarkdownParseOptions = {},
): MarkdownParseResult {
  const { body, frontMatter: parsedFrontMatter } = extractFrontMatter(markdown);
  const frontMatter = { ...parsedFrontMatter, ...options.metadata };
  const documentType = resolveDocumentType(options.documentType, frontMatter.type);
  const items = tokenize(body);
  const firstHeadingIndex = items.findIndex(item => item.kind === 'heading' && item.level === 1);
  const headingTitle = firstHeadingIndex >= 0
    ? (items[firstHeadingIndex] as HeadingItem).text
    : undefined;
  if (firstHeadingIndex >= 0) items.splice(firstHeadingIndex, 1);

  const title = stringValue(frontMatter, 'title')
    ?? headingTitle
    ?? options.sourceName
    ?? 'Untitled Document';

  const document = documentType === 'journal'
    ? createJournalDocument(title, items, frontMatter)
    : documentType === 'official'
      ? createOfficialDocument(title, items, frontMatter)
      : createThesisDocument(title, items, frontMatter);

  return { document, frontMatter };
}

export function parseMarkdownBlocks(markdown: string): ContentBlock[] {
  const { body } = extractFrontMatter(markdown);
  return tokenize(body).map(item => item.kind === 'heading'
    ? ({
        type: `heading${Math.min(item.level, 4)}`,
        text: item.text,
      } as ContentBlock)
    : item.block);
}

function extractFrontMatter(markdown: string): { body: string; frontMatter: MarkdownFrontMatter } {
  const normalized = markdown.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  const lines = normalized.split('\n');
  if (lines[0]?.trim() !== '---') return { body: normalized, frontMatter: {} };

  const closingIndex = lines.findIndex((line, index) => index > 0 && line.trim() === '---');
  if (closingIndex < 0) return { body: normalized, frontMatter: {} };

  const frontMatter: MarkdownFrontMatter = {};
  for (let index = 1; index < closingIndex; index += 1) {
    const match = /^([A-Za-z][\w-]*):\s*(.*)$/.exec(lines[index]);
    if (!match) continue;
    const key = match[1];
    const rawValue = match[2].trim();
    if (rawValue === '|' || rawValue === '>') {
      const parts: string[] = [];
      while (index + 1 < closingIndex && /^\s+/.test(lines[index + 1])) {
        index += 1;
        parts.push(lines[index].trim());
      }
      frontMatter[key] = rawValue === '|' ? parts.join('\n') : parts.join(' ');
    } else {
      frontMatter[key] = parseFrontMatterValue(rawValue);
    }
  }

  return { body: lines.slice(closingIndex + 1).join('\n'), frontMatter };
}

function parseFrontMatterValue(value: string): FrontMatterValue {
  const trimmed = value.trim();
  const unquoted = stripQuotes(trimmed);
  if (unquoted !== trimmed) return unquoted;
  if (/^\[.*\]$/.test(unquoted)) {
    const inner = unquoted.slice(1, -1).trim();
    return inner ? splitList(inner) : [];
  }
  if (/^(true|false)$/i.test(unquoted)) return unquoted.toLowerCase() === 'true';
  if (/^-?\d+(?:\.\d+)?$/.test(unquoted)) return Number(unquoted);
  return unquoted;
}

function tokenize(markdown: string): MarkdownItem[] {
  const lines = markdown.replace(/\r\n?/g, '\n').split('\n');
  const items: MarkdownItem[] = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();
    if (!trimmed) {
      index += 1;
      continue;
    }

    const heading = /^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (heading) {
      items.push({ kind: 'heading', level: heading[1].length, text: cleanInline(heading[2]) });
      index += 1;
      continue;
    }

    const fence = /^\s*(```+|~~~+)\s*([\w.+-]*)\s*$/.exec(line);
    if (fence) {
      const code: string[] = [];
      const fenceMarker = fence[1][0];
      index += 1;
      while (index < lines.length && !new RegExp(`^\\s*${escapeRegExp(fenceMarker)}{${fence[1].length},}\\s*$`).test(lines[index])) {
        code.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      items.push({ kind: 'block', block: { type: 'code_block', text: code.join('\n'), language: fence[2] || undefined } });
      continue;
    }

    if (trimmed === '$$' || (trimmed.startsWith('$$') && !trimmed.endsWith('$$'))) {
      const equation: string[] = [];
      if (trimmed !== '$$') equation.push(trimmed.slice(2));
      index += 1;
      while (index < lines.length && !lines[index].trim().endsWith('$$')) {
        equation.push(lines[index].trim());
        index += 1;
      }
      if (index < lines.length) {
        equation.push(lines[index].trim().replace(/\$\$$/, ''));
        index += 1;
      }
      items.push({ kind: 'block', block: { type: 'equation', latex: equation.join(' ').trim() } });
      continue;
    }
    if (/^\$\$[\s\S]*\$\$$/.test(trimmed)) {
      items.push({ kind: 'block', block: { type: 'equation', latex: trimmed.slice(2, -2).trim() } });
      index += 1;
      continue;
    }

    if (/^<!--\s*(?:page[-_ ]?break|newpage)\s*-->$/i.test(trimmed)) {
      items.push({ kind: 'block', block: { type: 'page_break' } });
      index += 1;
      continue;
    }

    if (/^ {0,3}((\*|-|_)\s*){3,}$/.test(line)) {
      items.push({ kind: 'block', block: { type: 'horizontal_rule' } });
      index += 1;
      continue;
    }

    const image = /^!\[([^\]]*)\]\((.+?)(?:\s+["']([^"']*)["'])?\)\s*$/.exec(trimmed);
    if (image) {
      const path = stripAngleBrackets(image[2].trim());
      const fallbackCaption = path.split(/[\\/]/).pop() ?? 'Figure';
      items.push({
        kind: 'block',
        block: { type: 'figure', path, caption: cleanInline(image[1] || image[3] || fallbackCaption) },
      });
      index += 1;
      continue;
    }

    if (isTableStart(lines, index)) {
      const headers = splitTableRow(lines[index]).map(cleanInline);
      const data: string[][] = [];
      index += 2;
      while (index < lines.length && lines[index].includes('|') && lines[index].trim()) {
        const row = splitTableRow(lines[index]).map(cleanInline);
        data.push(headers.map((_, cellIndex) => row[cellIndex] ?? ''));
        index += 1;
      }
      items.push({ kind: 'block', block: { type: 'table', caption: '', headers, data } });
      continue;
    }

    if (/^\s*>/.test(line)) {
      const quote: string[] = [];
      while (index < lines.length && /^\s*>/.test(lines[index])) {
        quote.push(lines[index].replace(/^\s*>\s?/, ''));
        index += 1;
      }
      items.push({ kind: 'block', block: { type: 'blockquote', text: cleanInline(joinSoftLines(quote)) } });
      continue;
    }

    const listMatch = /^(\s*)([-+*]|\d+[.)])\s+(.+)$/.exec(line);
    if (listMatch) {
      while (index < lines.length) {
        const item = /^(\s*)([-+*]|\d+[.)])\s+(.+)$/.exec(lines[index]);
        if (!item) break;
        const indentation = item[1].replace(/\t/g, '    ').length;
        items.push({
          kind: 'block',
          block: {
            type: 'list_item',
            text: cleanInline(item[3]),
            level: Math.floor(indentation / 2),
            ordered: /^\d/.test(item[2]),
            marker: item[2],
          },
        });
        index += 1;
      }
      continue;
    }

    const paragraph: string[] = [];
    while (index < lines.length && lines[index].trim() && !startsSpecialBlock(lines, index)) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    if (paragraph.length === 0) {
      paragraph.push(trimmed);
      index += 1;
    }
    items.push({ kind: 'block', block: { type: 'paragraph', text: cleanInline(joinSoftLines(paragraph)) } });
  }

  return items;
}

function startsSpecialBlock(lines: string[], index: number): boolean {
  const line = lines[index];
  const trimmed = line.trim();
  return /^ {0,3}#{1,6}\s+/.test(line)
    || /^\s*(```+|~~~+)/.test(line)
    || trimmed.startsWith('$$')
    || /^<!--\s*(?:page[-_ ]?break|newpage)\s*-->$/i.test(trimmed)
    || /^ {0,3}((\*|-|_)\s*){3,}$/.test(line)
    || /^!\[[^\]]*\]\(.+\)\s*$/.test(trimmed)
    || /^\s*>/.test(line)
    || /^(\s*)([-+*]|\d+[.)])\s+/.test(line)
    || isTableStart(lines, index);
}

function isTableStart(lines: string[], index: number): boolean {
  if (index + 1 >= lines.length || !lines[index].includes('|')) return false;
  const separator = lines[index + 1].trim();
  return /^\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?$/.test(separator);
}

function splitTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  const cells: string[] = [];
  let current = '';
  let escaped = false;
  for (const character of trimmed) {
    if (escaped) {
      current += character;
      escaped = false;
    } else if (character === '\\') {
      escaped = true;
    } else if (character === '|') {
      cells.push(current.trim());
      current = '';
    } else {
      current += character;
    }
  }
  cells.push(current.trim());
  return cells;
}

function createThesisDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): ThesisDocument {
  const author = stringValue(metadata, 'author');
  const department = stringValue(metadata, 'department', 'institution');
  const date = stringValue(metadata, 'date');
  const cover: ContentBlock[] = [
    { type: 'spacer', lines: 4 },
    { type: 'centered_text', text: title, font_size_pt: 22, bold: true },
  ];
  if (author) cover.push({ type: 'centered_text', text: author, font_size_pt: 14 });
  if (department) cover.push({ type: 'centered_text', text: department, font_size_pt: 14 });
  if (date) cover.push({ type: 'centered_text', text: date, font_size_pt: 14 });
  cover.push({ type: 'page_break' });

  return {
    type: 'thesis',
    meta: {
      title,
      titleEn: stringValue(metadata, 'titleEn', 'title_en'),
      author,
      supervisor: stringValue(metadata, 'supervisor'),
      department,
      major: stringValue(metadata, 'major'),
      degree: degreeValue(metadata.degree),
      date,
      keywords: arrayValue(metadata, 'keywords'),
      abstract: stringValue(metadata, 'abstract'),
      abstractEn: stringValue(metadata, 'abstractEn', 'abstract_en'),
      studentId: stringValue(metadata, 'studentId', 'student_id'),
    },
    cover,
    sections: buildSectionTree(items).map((node, index) => thesisSection(node, `${index + 1}`)),
  };
}

function createJournalDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): JournalArticle {
  const authorNames = arrayValue(metadata, 'authors', 'author') ?? [];
  const affiliations = arrayValue(metadata, 'affiliations', 'affiliation', 'institution') ?? [];
  const authors: JournalAuthor[] = authorNames.map(name => ({
    name,
    affiliations: affiliations.map(institution => ({ institution })),
  }));

  return {
    type: 'journal',
    meta: {
      title,
      runningTitle: stringValue(metadata, 'runningTitle', 'running_title'),
      authors,
      journalName: stringValue(metadata, 'journalName', 'journal'),
      articleType: articleTypeValue(metadata.articleType ?? metadata.article_type),
      abstract: stringValue(metadata, 'abstract'),
      keywords: arrayValue(metadata, 'keywords'),
      doi: stringValue(metadata, 'doi'),
    },
    sections: buildSectionTree(items).map((node, index) => journalSection(node, `${index + 1}`)),
  };
}

function createOfficialDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): OfficialDocument {
  const categoryValue = stringValue(metadata, 'documentCategory', 'category');
  const documentCategory = OFFICIAL_CATEGORIES.includes(categoryValue as OfficialDocCategory)
    ? categoryValue as OfficialDocCategory
    : '通知';
  const body = items.map(item => item.kind === 'heading'
    ? ({ type: `heading${Math.min(item.level, 4)}`, text: item.text } as ContentBlock)
    : item.block);

  return {
    type: 'official',
    meta: {
      title,
      issuingAuthority: stringValue(metadata, 'issuingAuthority', 'authority', 'organization') ?? '',
      documentNumber: stringValue(metadata, 'documentNumber', 'document_number') ?? '',
      documentCategory,
      primaryRecipients: arrayValue(metadata, 'primaryRecipients', 'recipients') ?? [],
      ccRecipients: arrayValue(metadata, 'ccRecipients', 'cc'),
      signatureAuthority: stringValue(metadata, 'signatureAuthority', 'signature_authority'),
      date: stringValue(metadata, 'date') ?? '',
      urgency: urgencyValue(metadata.urgency),
    },
    body,
  };
}

function buildSectionTree(items: MarkdownItem[]): SectionNode[] {
  const roots: SectionNode[] = [];
  const stack: SectionNode[] = [];
  let preamble: ContentBlock[] = [];

  for (const item of items) {
    if (item.kind === 'block') {
      if (stack.length > 0) stack[stack.length - 1].content.push(item.block);
      else preamble.push(item.block);
      continue;
    }

    const node: SectionNode = { level: item.level, title: item.text, content: [], children: [] };
    while (stack.length > 0 && stack[stack.length - 1].level >= item.level) stack.pop();
    if (stack.length > 0) stack[stack.length - 1].children.push(node);
    else roots.push(node);
    stack.push(node);
  }

  if (preamble.length > 0) {
    roots.unshift({ level: 1, title: '正文', content: preamble, children: [] });
  }
  if (roots.length === 0) {
    roots.push({ level: 1, title: '正文', content: [], children: [] });
  }
  return roots;
}

function thesisSection(node: SectionNode, id: string): ThesisSection {
  return {
    id: `section-${id}-${slugify(node.title)}`,
    type: thesisSectionType(node.title),
    title: node.title,
    content: node.content,
    subsections: node.children.map((child, index) => thesisSection(child, `${id}-${index + 1}`)),
  };
}

function journalSection(node: SectionNode, id: string): JournalSection {
  return {
    id: `section-${id}-${slugify(node.title)}`,
    type: journalSectionType(node.title),
    title: node.title,
    content: node.content,
    subsections: node.children.map((child, index) => journalSection(child, `${id}-${index + 1}`)),
  };
}

function thesisSectionType(title: string): ThesisSection['type'] {
  const normalized = title.toLowerCase();
  if (/abstract|摘要/.test(normalized)) return 'abstract';
  if (/acknowledg|致谢/.test(normalized)) return 'acknowledgement';
  if (/references|bibliography|参考文献/.test(normalized)) return 'references';
  if (/appendix|附录/.test(normalized)) return 'appendix';
  if (/contents|目录/.test(normalized)) return 'toc';
  return 'chapter';
}

function journalSectionType(title: string): JournalSection['type'] {
  const normalized = title.toLowerCase();
  if (/method|方法/.test(normalized)) return 'methods';
  if (/material|材料/.test(normalized)) return 'materials';
  if (/result|结果/.test(normalized)) return 'results';
  if (/discussion|讨论/.test(normalized)) return 'discussion';
  if (/conclusion|结论/.test(normalized)) return 'conclusion';
  if (/background|背景/.test(normalized)) return 'background';
  if (/related|相关工作|文献综述/.test(normalized)) return 'related_work';
  if (/experiment|实验/.test(normalized)) return 'experiment';
  if (/analysis|分析/.test(normalized)) return 'analysis';
  if (/appendix|附录/.test(normalized)) return 'appendix';
  return 'introduction';
}

function resolveDocumentType(option: DocumentType | undefined, value: FrontMatterValue | undefined): DocumentType {
  if (option) return option;
  const candidate = typeof value === 'string' ? value : 'thesis';
  if (!DOCUMENT_TYPES.includes(candidate as DocumentType)) {
    throw new Error(`Unsupported Markdown document type: ${candidate}`);
  }
  return candidate as DocumentType;
}

function stringValue(metadata: MarkdownFrontMatter, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = metadata[key];
    if (typeof value === 'string' && value.trim()) return value.trim();
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  }
  return undefined;
}

function arrayValue(metadata: MarkdownFrontMatter, ...keys: string[]): string[] | undefined {
  for (const key of keys) {
    const value = metadata[key];
    if (Array.isArray(value)) return value.map(item => item.trim()).filter(Boolean);
    if (typeof value === 'string' && value.trim()) return splitList(value);
  }
  return undefined;
}

function splitList(value: string): string[] {
  return value.split(/\s*[,;；]\s*/).map(stripQuotes).map(item => item.trim()).filter(Boolean);
}

function degreeValue(value: FrontMatterValue | undefined): ThesisDocument['meta']['degree'] {
  return value === 'bachelor' || value === 'master' || value === 'doctor' ? value : undefined;
}

function articleTypeValue(value: FrontMatterValue | undefined): JournalArticle['meta']['articleType'] {
  const allowed = ['research', 'review', 'short_communication', 'case_study', 'letter'] as const;
  return typeof value === 'string' && allowed.includes(value as typeof allowed[number])
    ? value as typeof allowed[number]
    : undefined;
}

function urgencyValue(value: FrontMatterValue | undefined): OfficialDocument['meta']['urgency'] {
  return value === 'normal' || value === 'urgent' || value === 'most_urgent' ? value : undefined;
}

function cleanInline(text: string): string {
  return text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/!\[([^\]]*)\]\([^)]+\)/g, '$1')
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/(\*\*|__)(.*?)\1/g, '$2')
    .replace(/(\*|_)(.*?)\1/g, '$2')
    .replace(/~~(.*?)~~/g, '$1')
    .replace(/<[^>]+>/g, '')
    .replace(/\\([\\`*{}\[\]()#+.!_>-])/g, '$1')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .trim();
}

function joinSoftLines(lines: string[]): string {
  return lines.reduce((result, line) => {
    const current = line.trim();
    if (!result) return current;
    return result + (endsWithCjk(result) && startsWithCjk(current) ? '' : ' ') + current;
  }, '');
}

function startsWithCjk(value: string): boolean {
  return /^[\u3000-\u9fff]/u.test(value);
}

function endsWithCjk(value: string): boolean {
  return /[\u3000-\u9fff]$/u.test(value);
}

function stripQuotes(value: string): string {
  return value.replace(/^(?:"([\s\S]*)"|'([\s\S]*)')$/, '$1$2');
}

function stripAngleBrackets(value: string): string {
  return value.startsWith('<') && value.endsWith('>') ? value.slice(1, -1) : value;
}

function slugify(value: string): string {
  const slug = value.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '-').replace(/^-|-$/g, '');
  return slug || 'untitled';
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
