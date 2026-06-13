# OpenThesis — AI-powered Document Template Engine

## Project Overview
Parse any .docx template (thesis, journal article, or official document), write content in structured JSON, export a submission-ready DOCX. The same engine handles all three document types — the template parser and renderer are document-type-agnostic.

## Architecture

```
packages/
  document-schema/     → Domain models: Thesis, JournalArticle, OfficialDocument
  template-parser/     → .docx → JSON style DSL (JSZip + fast-xml-parser)
  docx-renderer/       → JSON content + template → .docx (dolanmiu/docx)
  equation-engine/     → LaTeX → Plain Text (V1) / OMML (V3)
  cli/                 → `thesis parse|build|init` commands

examples/
  sample-thesis.json   → Sample thesis content
```

## Three Document Types

| Type | CLI flag | Standards |
|------|----------|-----------|
| Thesis (学位论文) | `--type thesis` | University-specific formatting |
| Journal Article (期刊论文) | `--type journal` | Elsevier, Springer, IEEE, etc. |
| Official Document (公文) | `--type official` | GB/T 9704-2012 (《党政机关公文格式》) |

## Key Design Decisions

### 1. Template-Driven, Not Hardcoded
All formatting parameters come from the parsed `DocumentTemplate` object. The `createUSTBTemplate()` function provides a fallback.

### 2. Document-Type-Agnostic Core
The template-parser and docx-renderer don't care what kind of document they're processing. Only the document-schema types and CLI commands are document-type-specific.

### 3. Style Inheritance Resolution
DOCX styles have a `basedOn` chain. The template-parser recursively resolves inheritance so each style's JSON is fully self-contained.

### 4. Semantic Role Detection (Heuristic V1)
Style roles are detected by regex matching on style names plus outline level fallback. In V4, this will be replaced by ML-based layout understanding.

### 5. Official Document Support (V1)
Complete block types for Chinese official documents: `red_header`, `document_number`, `recipient_line`, `signature_block`, `attachment_note`.

## Commands

```bash
pnpm install
pnpm build

# Parse templates
thesis parse <template.docx> --type thesis|journal|official --org <name>

# Create samples
thesis init --type thesis    # → thesis-content.json
thesis init --type journal   # → journal-content.json
thesis init --type official  # → official-content.json

# Build documents
thesis build <content.json> -t <template.json> -o output.docx
```

## Current Limitations

| Area | Limitation | Plan |
|------|-----------|------|
| Style role detection | Regex-based, fails on auto-numbered styles | V4: ML/layout-based |
| Equation rendering | Plain text only, no OMML | V3: LaTeX → MathML → OMML |
| Image embedding | Placeholder text only | V2: Read + embed image files |
| Template parsing | Works best with real Word templates | Current: works for most templates |
| Column support | Single-column only | V2: Double-column for journals |

## Next Steps
1. Get real university .docx templates (created in Word) to test parser quality
2. Add image embedding to docx-renderer
3. Add OMML equation generation (LaTeX → MathML → OMML)
4. Add Markdown → JSON content parser
5. Template marketplace (community-contributed templates)
