# Markdown import contract

OpenThesis accepts UTF-8 `.md` and `.markdown` files. It intentionally supports a predictable manuscript subset rather than every Markdown extension.

## Front matter

Place YAML-style scalar or inline-array metadata at the top:

```markdown
---
type: thesis
title: A Study of Open Documents
author: Jane Doe
degree: doctor
keywords: [documents, reproducibility]
---
```

Shared fields: `type`, `title`, `date`, `keywords`, `abstract`.

Thesis fields: `titleEn`, `author`, `supervisor`, `department`, `major`, `degree`, `abstractEn`, `studentId`.

Journal fields: `authors`, `affiliation` or `affiliations`, `journalName`, `articleType`, `doi`. Inline author arrays create authors with the listed shared affiliations.

Official-document fields: `issuingAuthority` (aliases `authority`, `organization`), `documentNumber`, `documentCategory` (alias `category`), `primaryRecipients` (alias `recipients`), `ccRecipients`, `signatureAuthority`, `date`, `urgency`.

The `type` value is `thesis`, `journal`, or `official`. The CLI `--type` option overrides front matter when explicitly supplied.

## Supported body syntax

| Markdown | OpenThesis block |
| --- | --- |
| `#` through `######` | Nested sections or heading blocks |
| Paragraphs | `paragraph` |
| `- item` / `1. item` | `list_item` with nesting and order |
| Fenced code | `code_block` with language |
| `> quote` | `blockquote` |
| `$$...$$` | `equation` |
| `![caption](path)` | `figure` |
| Pipe tables | `table` |
| `---` | `horizontal_rule` |
| `<!-- pagebreak -->` | `page_break` |

The first level-one heading is treated as the document title and is not duplicated in the body. Front-matter `title` takes precedence over its text. Remaining headings become nested thesis or journal sections; official-document headings remain body heading blocks.

Inline emphasis, links, and code are converted to plain text because the current content schema stores paragraph text without inline Markdown ranges.

## Media paths

During direct Markdown builds, relative image paths resolve from the Markdown file's directory. After importing to JSON, paths resolve from the JSON file's directory.
