# @openthesis/markdown-parser

Converts a practical, documented subset of Markdown into the typed OpenThesis
document schema. It supports YAML-style front matter, nested headings,
paragraphs, ordered and unordered lists, blockquotes, fenced code, tables,
images, display equations, horizontal rules, and page breaks.

```ts
import { parseMarkdown } from '@openthesis/markdown-parser';

const document = parseMarkdown(markdown, {
  documentType: 'thesis',
  sourceName: 'my-thesis',
});
```

See the root README for front-matter fields and CLI usage.
