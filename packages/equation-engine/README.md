# @openthesis/equation-engine

Renderer-neutral LaTeX math parsing for OpenThesis. The package converts a practical thesis-oriented LaTeX subset into a typed AST; `@openthesis/docx-renderer` maps that AST to editable native Office Math (OMML).

## Supported native subset

| Category | Syntax examples |
| --- | --- |
| Fractions and roots | `\frac{a}{b}`, `\dfrac{a}{b}`, `\sqrt{x}`, `\sqrt[3]{x}` |
| Scripts | `x_i`, `x^{2}`, `x_i^{2}` |
| N-ary operators | `\sum`, `\prod`, `\coprod`, `\bigcap`, `\bigcup`, `\int`, `\iint`, `\iiint` with limits |
| Functions | `\sin`, `\cos`, `\tan`, `\log`, `\ln`, `\exp`, `\det`, `\operatorname{...}` |
| Limits | `\lim`, `\min`, `\max`, `\inf`, `\sup` with lower/upper annotations |
| Accents and bars | `\hat`, `\widehat`, `\bar`, `\vec`, `\dot`, `\ddot`, `\tilde`, `\overline`, `\underline` |
| Delimiters | Parentheses, square brackets, and `\left ... \right` with braces, angle brackets, vertical bars, or invisible `.` delimiters |
| Annotations | `\overset{...}{...}`, `\underset{...}{...}` |
| Binomials | `\binom{n}{k}` |
| Environments | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases`, `aligned` |
| Symbols | Common Greek letters, relations, arrows, operators, and spacing commands |

Example:

```ts
import { latexToMathAst } from '@openthesis/equation-engine';

const ast = latexToMathAst(String.raw`
  f(x)=\begin{cases}
    x^2 & x \geq 0 \\
    -x  & x < 0
  \end{cases}
`);
```

## Architecture

The AST deliberately contains no `docx` classes. This keeps parsing independently testable and allows another renderer to consume the same tree. Native OMML construction lives in `@openthesis/docx-renderer`.

The older `latexToPlainText` and optional Pandoc-based `latexToOMML` APIs remain available as compatibility fallbacks. Normal OpenThesis DOCX builds use the native AST path.

## Current limitations

- User-defined `\newcommand` macros are not expanded.
- `array` column specifications and uncommon AMS environments are not interpreted.
- Equation alignment is represented with OMML matrix rows; per-column alignment controls are not yet exposed.
- Text/font styling commands currently preserve content but do not fully reproduce every LaTeX math alphabet.
- Unknown commands degrade to readable command text instead of executing arbitrary TeX.

Run the focused regression tests from the repository root:

```bash
pnpm build
node --test tests/equation-engine.test.mjs tests/docx-renderer.test.mjs
```
