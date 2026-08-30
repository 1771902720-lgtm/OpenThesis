# Structured content quick reference

Every content file is one of three tagged documents:

- Thesis: `{ "type": "thesis", "meta": {...}, "cover": [...], "sections": [...] }`
- Journal: `{ "type": "journal", "meta": {...}, "sections": [...] }`
- Official: `{ "type": "official", "meta": {...}, "body": [...] }`

Use the TypeScript definitions in `packages/document-schema/src/index.ts` as the authoritative contract.

## Common blocks

```json
{ "type": "paragraph", "text": "Body text" }
{ "type": "heading2", "text": "Nested heading" }
{ "type": "equation", "latex": "E = mc^2" }
{ "type": "figure", "path": "figures/result.png", "caption": "Experimental result" }
{ "type": "table", "caption": "Results", "headers": ["A", "B"], "data": [["1", "2"]] }
{ "type": "list_item", "text": "First item", "ordered": true, "level": 0 }
{ "type": "code_block", "language": "ts", "text": "const x = 1;" }
{ "type": "blockquote", "text": "Quoted text" }
{ "type": "horizontal_rule" }
{ "type": "page_break" }
```

Table rows may be shorter than the header list; missing cells render empty. If `columnWidths` is present, it must contain one positive number per header.

Figure paths may be absolute or relative to the content file. Supported detection includes PNG, JPEG, GIF, and BMP.

## Document-specific notes

- Thesis sections support recursive `subsections` and section types such as `chapter`, `abstract`, `references`, and `appendix`.
- Journal metadata requires an `authors` array; each author has an `affiliations` array, which may be empty.
- Official metadata requires `issuingAuthority`, `documentNumber`, `documentCategory`, `primaryRecipients`, and `date`. Use block types such as `recipient_line`, `signature_block`, and `attachment_note` when direct control is needed.

Legacy `{cover_blocks, body_blocks}` JSON remains readable but should not be generated for new work.
