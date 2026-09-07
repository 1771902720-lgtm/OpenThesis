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


Quoted front-matter values remain strings: use `studentId: "00042"` to preserve
leading zeros. Unquoted numbers and booleans retain their existing scalar parsing.
A fenced code block closes only with the same marker and at least as many markers
as its opening fence, so four-marker fences can contain three-marker examples.
