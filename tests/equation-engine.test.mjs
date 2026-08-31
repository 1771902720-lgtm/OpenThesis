import test from 'node:test';
import assert from 'node:assert/strict';
import { isLatexMath, latexToMathAst, latexToPlainText } from '../packages/equation-engine/dist/index.js';

test('converts common LaTeX symbols and scripts to Unicode', () => {
  assert.equal(latexToPlainText('E = mc^{2}'), 'E = mc²');
  assert.equal(latexToPlainText('\\alpha + \\beta \\leq \\infty'), 'α + β ≤ ∞');
  assert.equal(latexToPlainText('x_{max} = \\frac{a}{b}'), 'xₘₐₓ = (a)/(b)');
});

test('detects LaTeX while ignoring ordinary prose', () => {
  assert.equal(isLatexMath('\\frac{a}{b}'), true);
  assert.equal(isLatexMath('x^{2}'), true);
  assert.equal(isLatexMath('ordinary text'), false);
});

test('builds a native-math AST for fractions, radicals, scripts, sums, and integrals', () => {
  const nodes = latexToMathAst('\\frac{a_1}{\\sqrt{x^2}} + \\sum_{i=1}^{n} x_i + \\int_0^1 f(x)');
  assert.equal(nodes[0].type, 'fraction');
  assert.equal(nodes[0].numerator[0].type, 'script');
  assert.equal(nodes[0].denominator[0].type, 'radical');
  assert.ok(nodes.some(node => node.type === 'sum' && node.subScript && node.superScript));
  assert.ok(nodes.some(node => node.type === 'integral' && node.subScript && node.superScript));
});

test('parses advanced macros into native-math nodes', () => {
  const nodes = latexToMathAst(
    '\\hat{x}+\\overline{AB}+\\lim_{x\\to0}\\frac{\\sin(x)}{x}+\\binom{n}{k}+\\overset{*}{=}',
  );
  assert.ok(nodes.some(node => node.type === 'accent' && node.accent === '̂'));
  assert.ok(nodes.some(node => node.type === 'bar' && node.position === 'top'));
  assert.ok(nodes.some(node => node.type === 'limit' && node.name === 'lim' && node.subScript));
  const fraction = nodes.find(node => node.type === 'fraction');
  assert.equal(fraction?.numerator[0].type, 'function');
  const binomial = nodes.find(node => node.type === 'delimiter' && node.opening === '(');
  assert.equal(binomial?.children[0].type, 'fraction');
  assert.equal(binomial?.children[0].bar, false);
  assert.ok(nodes.some(node => node.type === 'overUnder' && node.over));

  const nary = latexToMathAst('\\prod_{i=1}^{n}x_i')[0];
  assert.equal(nary.type, 'nary');
  assert.equal(nary.operator, '∏');
  assert.ok(nary.subScript && nary.superScript);
});

test('parses matrix, cases, and aligned environments with cells and rows', () => {
  const matrix = latexToMathAst('\\begin{pmatrix}a & \\frac{b}{c} \\\\ d & e\\end{pmatrix}')[0];
  assert.equal(matrix.type, 'matrix');
  assert.equal(matrix.environment, 'pmatrix');
  assert.deepEqual([matrix.opening, matrix.closing], ['(', ')']);
  assert.equal(matrix.rows.length, 2);
  assert.equal(matrix.rows[0].length, 2);
  assert.equal(matrix.rows[0][1][0].type, 'fraction');

  const cases = latexToMathAst('\\begin{cases}x^2 & x\\geq0 \\\\ -x & x<0\\end{cases}')[0];
  assert.equal(cases.type, 'matrix');
  assert.equal(cases.environment, 'cases');
  assert.deepEqual([cases.opening, cases.closing], ['{', '']);

  const aligned = latexToMathAst('\\begin{aligned}a&=b+c \\\\ d&=e\\end{aligned}')[0];
  assert.equal(aligned.type, 'matrix');
  assert.equal(aligned.rows.length, 2);
});

test('preserves scalable left/right delimiters in the AST', () => {
  const node = latexToMathAst('\\left\\langle \\frac{x}{y} \\right\\rangle')[0];
  assert.equal(node.type, 'delimiter');
  assert.deepEqual([node.opening, node.closing], ['⟨', '⟩']);
  assert.ok(node.children.some(child => child.type === 'fraction'));
});

test('keeps scripts on function arguments distinct from scripts on function names', () => {
  const argumentScript = latexToMathAst('\\sin x^2')[0];
  assert.equal(argumentScript.type, 'function');
  assert.equal(argumentScript.children[0].type, 'script');

  const functionScript = latexToMathAst('\\sin^2 x')[0];
  assert.equal(functionScript.type, 'script');
  assert.equal(functionScript.base[0].type, 'function');
  assert.deepEqual(functionScript.superScript, [{ type: 'run', text: '2' }]);
});
