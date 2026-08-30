import test from 'node:test';
import assert from 'node:assert/strict';
import { isLatexMath, latexToPlainText } from '../packages/equation-engine/dist/index.js';

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
