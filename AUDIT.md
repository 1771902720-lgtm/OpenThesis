# OpenThesis — Engineering Audit

Date: 2026-09-28
Scope: `packages/*`, `tests/*`, CLI, agent skill, repository metadata
Method: full source read of all six packages, plus executable probes rendered against the built `dist/` — every defect below was reproduced, not inferred.

Baseline at audit time: `pnpm check` green, 29 tests passing, `pnpm audit --prod` **failing**.

---

## 1. Headline findings

Four problems made the engine produce wrong output on real templates. All four are
now fixed and covered by regression tests.

### 1.1 The built-in template produced headings identical to body text

`assets/ustb-thesis-template.json` resolved **all 72 styles to one identical
signature** — `Times New Roman / 宋体 / 12pt / justified / no indent`. Levels 1–3
of a nested heading rendered exactly like the surrounding body text, and level 4
rendered as 16pt 仿宋 (the *official-document* body font).

Two independent causes:

1. **`template-parser` read only style-level `w:rPr`** and ignored `w:pPr/w:rPr`,
   which is where Chinese Word/WPS templates put 黑体/三号 for headings.
2. **`rawToParagraphStyle` invented values** (`|| 'Times New Roman'`, `|| '宋体'`,
   `|| 24`, `?? 0`, `mapAlignment(undefined) → 'justified'`), so "the template is
   silent" became indistinguishable from "the template explicitly asks for 12pt".

The second cause is the important one: it is why the renderer could never apply a
sensible default. Verified against a realistic fixture, the parser itself was
correct — the failure was in what it emitted and what the renderer then trusted.

### 1.2 Tables overflowed the page

`renderTable` hardcoded `totalWidth = 8504` (an A4 constant) while the table
declared `<w:tblW w:type="pct" w:w="100%">` — contradictory OOXML.

Reproduced with an 8391-twip page and 1701-twip side margins (printable 4989):
cells still summed to **8504 twips, a 70% overflow** past the right margin.

### 1.3 Tables ignored the template entirely

Cell and caption runs hardcoded `Times New Roman` / 黑体 / 宋体 at 21 half-points.
A template requesting Georgia / 楷体 / 36 half-points was silently overruled; the
`_template` parameter was accepted and never read.

### 1.4 The repository's own CI was failing

`adm-zip@0.6.0` (a production dependency of `@openthesis/equation-engine`) carries
one **high** and one **moderate** advisory, both patched in `0.6.1`.

`pnpm audit --prod --audit-level high` is the last step of `.github/workflows/ci.yml`.
The last recorded CI run predates the advisory, so `main` is green on GitHub but
would fail on any push today.

---

## 2. Findings by severity

Status: **Fixed** = changed in this branch with a regression test.

### P0 — broken output or data loss

| # | Finding | Evidence | Status |
|---|---|---|---|
| P0-1 | `adm-zip@0.6.0` high + moderate advisories | `pnpm audit --prod` exit 1 | Fixed |
| P0-2 | `import a.md -o a.MD` overwrote the Markdown **source** with JSON — the guard was a case-sensitive string compare | reproduced | Fixed |
| P0-3 | `cleanInline` corrupted ordinary prose: `my_file_name` → `myfilename`, `2 * 3 * 4` → `2  3  4`, `if a < b > c then` → `if a  c then`; `<[^>]+>` also deleted every `<https://…>` autolink | reproduced | Fixed |
| P0-4 | `\text{if } x>0` rendered as `ifx>0` — whitespace was consumed before the script check and never emitted | reproduced | Fixed |

### P1 — wrong on real inputs

| # | Finding | Status |
|---|---|---|
| P1-1 | `heading4` absent from the fallback table → 4th-level headings resolved to the body paragraph style | Fixed |
| P1-2 | Table width hardcoded to 8504 twips, contradicting its own 100% declaration | Fixed |
| P1-3 | Table typography hardcoded; template styles ignored | Fixed |
| P1-4 | Thesis and 公文 shared one fallback table, so a thesis without a template body rendered as 仿宋 三号 | Fixed |
| P1-5 | Parser ignored `w:pPr/w:rPr` | Fixed |
| P1-6 | `w:firstLineChars` / `leftChars` / `rightChars` are hundredths of a *character* but were stored as twips (`200` instead of `480`) | Fixed |
| P1-7 | Unanchored `/^(表\|Table\|题注)/` matched Word's built-in `TableGrid` and `Table of Contents`, turning them into level-3 headings | Fixed |
| P1-8 | Role detection ignored `w:type`, so character (`uChar`, `Hyperlink`), table and numbering styles were all assigned `paragraph`; `resolveStyle` then picked whichever came first in the XML | Fixed |
| P1-9 | No pattern ever produced the `equation` role, so a template's equation style (`MTConvertedEquation`) could never be used | Fixed |
| P1-10 | `mathComponents` had no `default` case → an unhandled AST node put `undefined` into the OMML children array | Fixed |
| P1-11 | Equation parse failures were swallowed by a bare `catch` with no diagnostic | Fixed |
| P1-12 | `--flag=value` was silently ignored: `import x.md --type=official` exited 0 and produced `"type":"thesis"`; `init --type=journal` wrote `thesis-content.json` | Fixed |
| P1-13 | Unknown flags were silently ignored: `build x.json --out zzz.docx` wrote `x.docx` | Fixed |
| P1-14 | UTF-8-BOM JSON input failed with `Unexpected token '\uFEFF'` | Fixed |
| P1-15 | Front-matter block sequences (`keywords:` + `- item`) were skipped and the key became `""` | Fixed |
| P1-16 | Quoted list elements were split before quote handling: `["a, b", c]` → `['"a','b"','c']` | Fixed |
| P1-17 | `# C#` lost its `#`; the closing-hash pattern did not require a preceding space | Fixed |
| P1-18 | Single-column tables were never detected (separator regex required a second pipe group) | Fixed |
| P1-19 | A fence with an extended info string (` ```js title=x `) failed to match, so the opener became a paragraph and the closing fence swallowed the rest of the file | Fixed |
| P1-20 | `latexToPlainText` used an ordered replace chain, so any command that prefixes another was corrupted: `\propto` → `Πto`, `\cdots` → `·s`, `\simeq` → `∼eq` | Fixed |
| P1-21 | Bare `x^2` / `x_1` were left as literal caret/underscore | Fixed |
| P1-22 | `[^}]+` cannot match nested braces, so `\frac{a_{1}}{b}` leaked as `\fraca_1/b` | Fixed |
| P1-23 | `/\\(left\|right)\s*/` also matched `\leftarrow`, so the real `\right` was eaten and the delimiter swallowed the rest of the equation | Fixed |

### P2 — correctness and maintainability, still open

| # | Finding |
|---|---|
| P2-1 | `latexToOMML` falls back to Unicode whenever pandoc is missing **or fails for any reason**, silently. The path is effectively dead — `docx-renderer` imports only `latexToMathAst` / `latexToPlainText`. |
| P2-2 | `\operatorname*{argmax}_{x}`: the `*` is not consumed, yielding an empty function name plus a stray run and group. |
| P2-3 | `\begin{aligned}…\end{matrix}` (mismatched name) silently consumes the rest of the input as the environment body. |
| P2-4 | Unknown commands lose their backslash in the AST (`\foo` → run `foo`). |
| P2-5 | `x^{a}^{b}` silently overwrites the first superscript. |
| P2-6 | Escaped `&` after a row break is mis-detected (`\\&`), merging two cells into one. |
| P2-7 | Only the body-level `w:sectPr` is read; paragraph-level section breaks are ignored, so multi-section templates (cover + body, landscape appendix) expose one section. |
| P2-8 | `removeNSPrefix: false` hardcodes the `w:` prefix; a styles.xml using another prefix yields a zero-style template with no error. |
| P2-9 | `basedOn` cycles are accepted and produce order-dependent merges. |
| P2-10 | `lineSpacing` stores the raw `w:line` and drops `lineRule`, so `auto` (240ths) and `exact` (twips) are conflated. |
| P2-11 | Markdown list nesting is `floor(indent / 2)`, so standard 2- and 4-space nesting is mis-levelled. |
| P2-12 | Inline formatting is destroyed on import while `InlineRange` / `RichParagraph` exist in the schema and are referenced nowhere — dead types advertising a capability that does not exist. |
| P2-13 | Lists render as plain text runs with a manual `•` / `1.` prefix, not Word numbering. README:164 claims "native document rendering". |
| P2-14 | No document validator exists anywhere; the only runtime checks live in the renderer. |
| P2-15 | `$$a=b$$ trailing` mis-parses and then consumes following lines until one ends with `$$`. |
| P2-16 | A document starting with `---` and no front matter silently drops everything up to the next `---`. |
| P2-17 | An invalid official `category` silently becomes `通知`; `degree: PhD` silently becomes `undefined`. |
| P2-18 | `resolveStyle` scans `styleRoles` linearly for every block — O(styles × blocks). |
| P2-19 | CLI: no `--version`, no stdin/stdout, no `mkdir -p` for the output directory, no overwrite protection for an existing output (only the input/template equality guards were added). |
| P2-20 | Agent-skill documentation mismatches: the undocumented `thesis` prefix (all skill examples omit it while `command = args[0]` requires it), overstated fence support, missing block-list caveat, `schema.md` omits `backMatter`, and `workflows.md` hardcodes a Linux `/root/.codex/...` path. |

### Repository hygiene

| # | Finding | Status |
|---|---|---|
| H-1 | 24 references to the pre-rename owner `1771902720-lgtm` across 13 files (CI badge, clone URLs, star-history chart, every `package.json` `repository`/`homepage`/`bugs`). GitHub redirects them, so they mostly work, but npm metadata should be canonical. | Fixed |
| H-2 | `drm-fix-report.json` — a 549-line personal research report (北京科技大学 动力松弛法) committed to the public repository root. | Removed |
| H-3 | `thesis-content.json` — a `thesis init` output artifact committed at the root. | Removed |
| H-4 | `.gitignore`'s blanket `*.docx` also blocked template fixtures, directly obstructing the roadmap item "get real university and journal templates to expand parser compatibility fixtures". | Fixed (`examples/` and `tests/fixtures/` are now exempt) |

---

## 3. What changed

| Package | Change |
|---|---|
| `document-schema` | `FontSettings.name/eastAsia/size` and `ParagraphFormatting.alignment` are optional — a parsed template declares only what it specifies. Added `StyleRole` (`BlockType` + `table_header` + `figure_caption`). |
| `template-parser` | Merge `w:pPr/w:rPr` under the style-level `w:rPr`. Convert `*Chars` indents to twips. Stop inventing fallback values. Skip non-paragraph styles when assigning block roles. Anchor the caption pattern; add an `equation` pattern. Order role assignment so an explicit or `w:default` style wins the renderer's first-match lookup. |
| `docx-renderer` | Split `THESIS_STYLES` from `OFFICIAL_STYLES` (GB/T 9704-2012) and add `heading4`, `table`, `table_header`, `figure_caption`. `resolveStyle` merges the template **per property** over the role default. Table width derives from `page.width − margins`; table and caption typography resolve from the template. Image width fits the printable page. Exhaustive `mathComponents` switch. Equation fallback now warns. |
| `equation-engine` | Exact-command symbol lookup (no prefix collisions). Brace-aware group matching for `\frac`, `\sqrt`, accents and `\text`. Bare-script conversion. `\text{}` keeps its spaces; the insignificant space after it is dropped. `\leftarrow` no longer matches `\left`. |
| `cli` | `--flag=value` support, unknown-flag rejection, case-insensitive input/output equality (plus output ≠ template), BOM stripping, `init` shares the same option reader. |
| `markdown-parser` | Emphasis requires real delimiters (no intraword `_`, no space-flanked `*`); only tag-shaped `<…>` is stripped and autolinks keep their target. Heading closing-hash requires a space. Front-matter block sequences. Quote-aware list splitting. Single-column tables. Extended fence info strings. |
| `assets/ustb-thesis-template.json` | The 72 entries that were byte-identical to the old invented defaults were stripped to empty objects. **This file still needs regenerating from the original `.docx`** (`thesis parse <template.docx>`) to recover any formatting the old parser dropped — see §5. |

Test count: **29 → 54**, all passing. `pnpm audit --prod --audit-level high` now exits 0.

---

## 4. Reproducing the original defects

The probes used during the audit are reproduced by the regression tests added in
`tests/`. To see the original behaviour, check out the parent commit and run:

```bash
pnpm install && pnpm build
pnpm audit --prod --audit-level high          # exits 1
node --test tests/*.test.mjs                  # 29 tests, all pass
```

Then render a four-level heading, or a table on a narrow page, and inspect
`word/document.xml` — the probes are described in §1.1–§1.3.

---

## 5. Known gap that needs your input

`assets/ustb-thesis-template.json` was generated by the **old** parser and still
contains no per-style formatting. The renderer now compensates with role defaults,
so output is correct, but the template file itself carries no information.

The original `.docx` is not in the repository — `meta.sourceFile` points at
`C:\Users\17719\Desktop\各类文件\学位论文 模板\《北京科技大学硕士学位论文模板》.docx`.
Re-running

```bash
node packages/cli/dist/index.js parse "《北京科技大学硕士学位论文模板》.docx" --type thesis --org "北京科技大学"
```

with the fixed parser may recover heading fonts and sizes that the `w:pPr/w:rPr`
bug was dropping. If the recovered file differs, commit it — it should improve
fidelity for every user of the built-in template.

---

## 6. Suggested next steps

1. Regenerate the USTB template JSON (§5).
2. Add real university and journal `.docx` files under `tests/fixtures/` — now
   unblocked by the `.gitignore` change — and assert parsed geometry and roles.
3. Close the remaining P2 equation gaps (`\operatorname*`, environment mismatch,
   unknown-command preservation) and surface unsupported LaTeX as a diagnostic
   instead of a silent Unicode downgrade.
4. Emit real Word numbering for lists instead of manual markers, or correct the
   README's "native document rendering" claim.
5. Add a document validator so a malformed content JSON fails with a clear message
   instead of deep inside the renderer.
