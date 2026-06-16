// ============================================================
// @openthesis/document-schema — Document Domain Models
// ============================================================
// Defines WHAT a document IS — independent of any file format,
// rendering engine, or template system.
//
// Three domain models, one style system:
//   1. Thesis (学位论文)             — 学士/硕士/博士论文
//   2. JournalArticle (期刊论文)      — 学术期刊投稿
//   3. OfficialDocument (公文)       — 党政机关公文 (GB/T 9704)
// ============================================================

// ============================================================
// ── Document Type Enum ─────────────────────────────────────
// ============================================================

export type DocumentType = 'thesis' | 'journal' | 'official';

// ============================================================
// ── TEMPLATE (parsed from .docx) ───────────────────────────
// ============================================================

export interface DocumentTemplate {
  meta: TemplateMeta;
  page: PageSettings;
  styles: Record<string, ParagraphStyle>;
  styleRoles: Record<string, BlockType>;
  styleInheritance?: Record<string, string>;
  documentType?: DocumentType;
}

export interface TemplateMeta {
  organization: string;
  name: string;
  documentType: DocumentType;
  subType?: string;
  sourceFile?: string;
  parserVersion: string;
  parsedAt: string;
}

export interface PageSettings {
  width: number;
  height: number;
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
  };
  headerDistance?: number;
  footerDistance?: number;
  columns?: number;
  columnGutter?: number;
}

// ── Styles (shared) ────────────────────────────────────────

export interface ParagraphStyle {
  font: FontSettings;
  paragraph: ParagraphFormatting;
  lineSpacing?: number;
}

export interface FontSettings {
  name: string;
  eastAsia: string;
  size: number;         // half-points (24 = 12pt)
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

export interface ParagraphFormatting {
  alignment: 'left' | 'center' | 'right' | 'justified' | 'distribute';
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  outlineLevel?: number;
}

// ════════════════════════════════════════════════════════════
// ── CONTENT BLOCKS ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════

export type BlockType =
  | 'heading1' | 'heading2' | 'heading3' | 'heading4'
  | 'paragraph' | 'paragraph_no_indent'
  | 'centered_text'
  | 'equation' | 'equation_numbered'
  | 'figure' | 'table'
  | 'list_item' | 'code_block' | 'blockquote'
  | 'horizontal_rule' | 'spacer' | 'page_break'
  // Official document blocks
  | 'red_header' | 'document_number' | 'recipient_line'
  | 'signature_block' | 'attachment_note';

export interface BaseBlock { type: BlockType; id?: string; }

export interface HeadingBlock extends BaseBlock {
  type: 'heading1' | 'heading2' | 'heading3' | 'heading4';
  text: string; number?: string;
}

export interface ParagraphBlock extends BaseBlock {
  type: 'paragraph' | 'paragraph_no_indent';
  text: string;
}

export interface CenteredTextBlock extends BaseBlock {
  type: 'centered_text';
  text: string; font_size_pt?: number; bold?: boolean;
}

export interface EquationBlock extends BaseBlock {
  type: 'equation' | 'equation_numbered';
  latex: string; number?: string;
}

export interface FigureBlock extends BaseBlock {
  type: 'figure';
  path: string; caption: string; number?: string;
  width?: number; height?: number;
}

export interface TableBlock extends BaseBlock {
  type: 'table';
  caption: string; headers: string[]; data: string[][];
  number?: string; columnWidths?: number[];
  showGridlines?: boolean; headerShading?: boolean;
}

export interface ListItemBlock extends BaseBlock {
  type: 'list_item';
  text: string; level?: number; ordered?: boolean; marker?: string;
}

export interface CodeBlockBlock extends BaseBlock {
  type: 'code_block'; text: string; language?: string;
}

export interface BlockquoteBlock extends BaseBlock {
  type: 'blockquote'; text: string;
}

export interface SpacerBlock extends BaseBlock { type: 'spacer'; lines: number; }
export interface PageBreakBlock extends BaseBlock { type: 'page_break'; }
export interface HorizontalRuleBlock extends BaseBlock { type: 'horizontal_rule'; }

// Official document blocks
export interface RedHeaderBlock extends BaseBlock {
  type: 'red_header'; text: string; font_size_pt?: number;
}
export interface DocumentNumberBlock extends BaseBlock {
  type: 'document_number'; text: string;
}
export interface RecipientLineBlock extends BaseBlock {
  type: 'recipient_line'; text: string; recipientType: 'primary' | 'cc';
}
export interface SignatureBlockBlock extends BaseBlock {
  type: 'signature_block'; authority: string; date: string;
}
export interface AttachmentNoteBlock extends BaseBlock {
  type: 'attachment_note'; attachments: string[];
}

export type ContentBlock =
  | HeadingBlock | ParagraphBlock | CenteredTextBlock
  | EquationBlock | FigureBlock | TableBlock
  | ListItemBlock | CodeBlockBlock | BlockquoteBlock
  | SpacerBlock | PageBreakBlock | HorizontalRuleBlock
  | RedHeaderBlock | DocumentNumberBlock | RecipientLineBlock
  | SignatureBlockBlock | AttachmentNoteBlock;

// ════════════════════════════════════════════════════════════
// ── 1. THESIS (学位论文) ────────────────────────────────────
// ════════════════════════════════════════════════════════════

export interface ThesisDocument {
  type: 'thesis';
  meta: ThesisMeta;
  cover: ContentBlock[];
  sections: ThesisSection[];
  backMatter?: ThesisBackMatter;
}

export interface ThesisMeta {
  title: string;
  titleEn?: string;
  author?: string;
  supervisor?: string;
  department?: string;
  major?: string;
  degree?: 'bachelor' | 'master' | 'doctor';
  date?: string;
  keywords?: string[];
  abstract?: string;
  abstractEn?: string;
  classificationNumber?: string;
  studentId?: string;
}

export interface ThesisSection {
  id: string;
  type: 'chapter' | 'abstract' | 'acknowledgement' | 'toc'
      | 'list_of_figures' | 'list_of_tables' | 'references' | 'appendix';
  title: string;
  number?: string;
  content: ContentBlock[];
  subsections?: ThesisSection[];
}

export interface ThesisBackMatter {
  references?: Reference[];
  appendices?: ThesisSection[];
  authorBiography?: string;
  declaration?: string;
  datasetInfo?: Record<string, string>;
}

// ════════════════════════════════════════════════════════════
// ── 2. JOURNAL ARTICLE (期刊论文) ──────────────────────────
// ════════════════════════════════════════════════════════════

export interface JournalArticle {
  type: 'journal';
  meta: JournalMeta;
  sections: JournalSection[];
  backMatter?: JournalBackMatter;
}

export interface JournalMeta {
  title: string;
  runningTitle?: string;
  authors: JournalAuthor[];
  journalName?: string;
  articleType?: 'research' | 'review' | 'short_communication' | 'case_study' | 'letter';
  abstract?: string;
  keywords?: string[];
  funding?: FundingInfo[];
  submittedAt?: string;
  acceptedAt?: string;
  doi?: string;
}

export interface JournalAuthor {
  name: string;
  nameLocal?: string;
  affiliations: Affiliation[];
  isCorresponding?: boolean;
  email?: string;
  orcid?: string;
  equalContribution?: boolean;
}

export interface Affiliation {
  institution: string;
  department?: string;
  city?: string;
  country?: string;
  postalCode?: string;
}

export interface FundingInfo {
  agency: string;
  grantNumber?: string;
}

export interface JournalSection {
  id: string;
  type: 'introduction' | 'methods' | 'results' | 'discussion'
      | 'conclusion' | 'materials' | 'background' | 'related_work'
      | 'experiment' | 'analysis' | 'appendix' | 'supplementary';
  title: string;
  number?: string;
  content: ContentBlock[];
  subsections?: JournalSection[];
}

export interface JournalBackMatter {
  references?: Reference[];
  supplementary?: string;
  authorContributions?: string;
  conflictOfInterest?: string;
  dataAvailability?: string;
  acknowledgments?: string;
}

export interface Reference {
  id: string;
  text: string;
  type?: 'journal' | 'book' | 'conference' | 'thesis' | 'web' | 'standard' | 'other';
  doi?: string;
  url?: string;
}

// ════════════════════════════════════════════════════════════
// ── 3. OFFICIAL DOCUMENT (公文 / GB/T 9704-2012) ───────────
// ════════════════════════════════════════════════════════════

export interface OfficialDocument {
  type: 'official';
  meta: OfficialMeta;
  body: ContentBlock[];
  attachments?: OfficialAttachment[];
}

export interface OfficialMeta {
  title: string;
  issuingAuthority: string;
  documentNumber: string;
  signer?: string;
  documentCategory: OfficialDocCategory;
  primaryRecipients: string[];
  ccRecipients?: string[];
  signatureAuthority?: string;
  date: string;
  notes?: string;
  issuingDepartment?: string;
  issueDate?: string;
  urgency?: 'normal' | 'urgent' | 'most_urgent';
  confidentiality?: 'unclassified' | 'secret' | 'confidential' | 'top_secret';
  distributionNumber?: string;
}

export type OfficialDocCategory =
  | '通知' | '通告' | '报告' | '请示' | '批复'
  | '函'   | '纪要' | '决定' | '意见' | '通报'
  | '议案' | '公告' | '命令' | '决议' | '公报';

export interface OfficialAttachment {
  order: number;
  title: string;
  content?: ContentBlock[];
}

// ════════════════════════════════════════════════════════════
// ── INLINE FORMATTING ──────────────────────────────────────
// ════════════════════════════════════════════════════════════

export interface InlineRange {
  offset: number; length: number;
  bold?: boolean; italic?: boolean;
  underline?: boolean; superscript?: boolean; subscript?: boolean;
  fontName?: string; fontSize?: number; color?: string;
}

export interface RichParagraph {
  text: string;
  ranges: InlineRange[];
}

// ════════════════════════════════════════════════════════════
// ── LEGACY FORMAT ──────────────────────────────────────────
// ════════════════════════════════════════════════════════════

export interface LegacyDocumentJSON {
  title: string;
  report_header?: string;
  cover_blocks: ContentBlock[];
  body_blocks: ContentBlock[];
}

export type OpenThesisDocument = ThesisDocument | JournalArticle | OfficialDocument;

