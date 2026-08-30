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
