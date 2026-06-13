// ============================================================
// @openthesis/equation-engine — LaTeX → Office Math Engine
// ============================================================
// V1: Plain-text rendering (matching existing behavior)
// V3: Full OMML generation for native Word equation editing
// ============================================================

/**
 * Render a LaTeX equation string to plain Unicode text.
 * This is the V1 fallback — preserves existing pipeline behavior.
 *
 * In V3, this will be replaced by:
 *   LaTeX → MathML → OMML (Office Math Markup Language)
 * which generates native editable equations in Word.
 *
 * @param latex — LaTeX equation source (e.g. "E = mc^2")
 * @returns Unicode plain-text representation
 */
export function latexToPlainText(latex: string): string {
  return latex
    // Basic Greek letters
    .replace(/\\alpha/g, 'α')
    .replace(/\\beta/g, 'β')
    .replace(/\\gamma/g, 'γ')
    .replace(/\\delta/g, 'δ')
    .replace(/\\Delta/g, 'Δ')
    .replace(/\\epsilon/g, 'ε')
    .replace(/\\zeta/g, 'ζ')
    .replace(/\\eta/g, 'η')
    .replace(/\\theta/g, 'θ')
    .replace(/\\Theta/g, 'Θ')
    .replace(/\\lambda/g, 'λ')
    .replace(/\\Lambda/g, 'Λ')
    .replace(/\\mu/g, 'μ')
    .replace(/\\nu/g, 'ν')
    .replace(/\\xi/g, 'ξ')
    .replace(/\\pi/g, 'π')
    .replace(/\\rho/g, 'ρ')
    .replace(/\\sigma/g, 'σ')
    .replace(/\\Sigma/g, 'Σ')
    .replace(/\\tau/g, 'τ')
    .replace(/\\phi/g, 'φ')
    .replace(/\\Phi/g, 'Φ')
    .replace(/\\psi/g, 'ψ')
    .replace(/\\Psi/g, 'Ψ')
    .replace(/\\omega/g, 'ω')
    .replace(/\\Omega/g, 'Ω')
    // Infinity
    .replace(/\\infty/g, '∞')
    // Basic operators
    .replace(/\\times/g, '×')
    .replace(/\\cdot/g, '·')
    .replace(/\\pm/g, '±')
    .replace(/\\mp/g, '∓')
    .replace(/\\leq/g, '≤')
    .replace(/\\geq/g, '≥')
    .replace(/\\neq/g, '≠')
    .replace(/\\approx/g, '≈')
    .replace(/\\equiv/g, '≡')
    .replace(/\\propto/g, '∝')
    .replace(/\\partial/g, '∂')
    .replace(/\\nabla/g, '∇')
    .replace(/\\int/g, '∫')
    .replace(/\\sum/g, 'Σ')
    .replace(/\\prod/g, 'Π')
    .replace(/\\sqrt/g, '√')
    // Superscript / subscript markers → Unicode
    .replace(/\^\{([^}]+)\}/g, (_, p1) => toSuperscript(p1))
    .replace(/_\{([^}]+)\}/g, (_, p1) => toSubscript(p1))
    // Clean up remaining LaTeX braces
    .replace(/[{}]/g, '')
    .trim();
}

/**
 * Future: Convert LaTeX to OMML XML string.
 * This will be the V3 implementation.
 *
 * Pipeline: LaTeX → MathML (via MathJax/temme) → OMML (via XSLT)
 *
 * @param latex — LaTeX equation source
 * @returns OMML XML string ready for insertion into DOCX
 */
export function latexToOMML(_latex: string): string {
  throw new Error(
    'OMML generation is planned for V3. ' +
    'Currently use latexToPlainText() for plain-text rendering, ' +
    'or contribute to the equation-engine to add MathML→OMML conversion.',
  );
}

/**
 * Check if a string looks like LaTeX math.
 */
export function isLatexMath(text: string): boolean {
  return /\\[a-zA-Z]+|\^\{|_\{|\\frac|\\sum|\\int|\\sqrt/.test(text);
}

// ── Unicode helpers ───────────────────────────────────────

const SUPERSCRIPTS: Record<string, string> = {
  '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
  '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
  '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
  'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
  'n': 'ⁿ', 'i': 'ⁱ', 'k': 'ᵏ', 'm': 'ᵐ', 't': 'ᵗ',
  'T': 'ᵀ',
};

const SUBSCRIPTS: Record<string, string> = {
  '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
  '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
  '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
  'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ', 'n': 'ₙ',
};

function toSuperscript(text: string): string {
  return [...text].map(c => SUPERSCRIPTS[c] || c).join('');
}

function toSubscript(text: string): string {
  return [...text].map(c => SUBSCRIPTS[c] || c).join('');
}
