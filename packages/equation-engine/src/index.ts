// ============================================================
// @openthesis/equation-engine — LaTeX → Office Math Engine
// ============================================================
// V3: LaTeX math AST for native DOCX OMML rendering.
// Keeps Unicode fallback and optional Pandoc extraction for compatibility.
// ============================================================

import { execFileSync } from 'child_process';
import { writeFileSync, unlinkSync, existsSync, mkdirSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { randomUUID } from 'crypto';
import AdmZip from 'adm-zip';

// ── Public API ────────────────────────────────────────────

export type LatexMathNode =
  | { type: 'run'; text: string }
  | { type: 'fraction'; numerator: LatexMathNode[]; denominator: LatexMathNode[] }
  | { type: 'radical'; children: LatexMathNode[]; degree?: LatexMathNode[] }
  | { type: 'script'; base: LatexMathNode[]; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] }
  | { type: 'sum' | 'integral'; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] };

/**
 * Parse a practical LaTeX math subset into a renderer-neutral tree.
 * The DOCX renderer maps this tree to native Office Math (OMML) elements.
 */
export function latexToMathAst(latex: string): LatexMathNode[] {
  return new LatexMathParser(latex).parse();
}

/**
 * Render a LaTeX equation string to OMML XML (native Word equation).
 * Uses pandoc for LaTeX → OMML conversion when available.
 * Falls back to Unicode plain-text when pandoc is not installed.
 */
export function latexToOMML(latex: string): string {
  // V2: Try pandoc OMML generation first
  const omml = generateOMMLviaPandoc(latex);
  if (omml) return omml;

  // Fallback: Unicode plain-text conversion
  return latexToPlainText(latex);
}

/**
 * Render a LaTeX equation to plain Unicode text (V1 fallback).
 * Handles Greek letters, superscripts/subscripts, common operators.
 */
export function latexToPlainText(latex: string): string {
  let result = latex
    // Greek letters (uppercase)
    .replace(/\\Gamma/g, 'Γ').replace(/\\Delta/g, 'Δ')
    .replace(/\\Theta/g, 'Θ').replace(/\\Lambda/g, 'Λ')
    .replace(/\\Xi/g, 'Ξ').replace(/\\Pi/g, 'Π')
    .replace(/\\Sigma/g, 'Σ').replace(/\\Upsilon/g, 'Υ')
    .replace(/\\Phi/g, 'Φ').replace(/\\Psi/g, 'Ψ')
    .replace(/\\Omega/g, 'Ω')
    // Greek letters (lowercase)
    .replace(/\\alpha/g, 'α').replace(/\\beta/g, 'β')
    .replace(/\\gamma/g, 'γ').replace(/\\delta/g, 'δ')
    .replace(/\\epsilon/g, 'ε').replace(/\\varepsilon/g, 'ε')
    .replace(/\\zeta/g, 'ζ').replace(/\\eta/g, 'η')
    .replace(/\\theta/g, 'θ').replace(/\\vartheta/g, 'ϑ')
    .replace(/\\iota/g, 'ι').replace(/\\kappa/g, 'κ')
    .replace(/\\lambda/g, 'λ').replace(/\\mu/g, 'μ')
    .replace(/\\nu/g, 'ν').replace(/\\xi/g, 'ξ')
    .replace(/\\pi/g, 'π').replace(/\\rho/g, 'ρ')
    .replace(/\\sigma/g, 'σ').replace(/\\tau/g, 'τ')
    .replace(/\\upsilon/g, 'υ').replace(/\\phi/g, 'φ')
    .replace(/\\varphi/g, 'φ').replace(/\\chi/g, 'χ')
    .replace(/\\psi/g, 'ψ').replace(/\\omega/g, 'ω')
    // Common operators & relations
    .replace(/\\infty/g, '∞').replace(/\\partial/g, '∂')
    .replace(/\\nabla/g, '∇').replace(/\\int/g, '∫')
    .replace(/\\sum/g, 'Σ').replace(/\\prod/g, 'Π')
    .replace(/\\sqrt/g, '√').replace(/\\propto/g, '∝')
    .replace(/\\times/g, '×').replace(/\\cdot/g, '·')
    .replace(/\\pm/g, '±').replace(/\\mp/g, '∓')
    .replace(/\\leq/g, '≤').replace(/\\geq/g, '≥')
    .replace(/\\neq/g, '≠').replace(/\\approx/g, '≈')
    .replace(/\\equiv/g, '≡').replace(/\\sim/g, '∼')
    .replace(/\\parallel/g, '∥').replace(/\\perp/g, '⊥')
    // Fractions, text styling
    .replace(/\\frac\{([^}]+)\}\{([^}]+)\}/g, '($1)/($2)')
    .replace(/\\text\{([^}]+)\}/g, '$1')
    .replace(/\\mathrm\{([^}]+)\}/g, '$1')
    .replace(/\\mathbf\{([^}]+)\}/g, '$1')
    .replace(/\\mathit\{([^}]+)\}/g, '$1')
    // Bars, hats, dots
    .replace(/\\bar\{([^}]+)\}/g, '$1̄')
    .replace(/\\hat\{([^}]+)\}/g, '$1̂')
    .replace(/\\dot\{([^}]+)\}/g, '$1̇')
    .replace(/\\ddot\{([^}]+)\}/g, '$1̈')
    .replace(/\\tilde\{([^}]+)\}/g, '$1̃')
    .replace(/\\vec\{([^}]+)\}/g, '$1⃗')
    // Norm: \|...\| → ‖...‖
    .replace(/\\\|/g, '‖')
    // Superscript with braces
    .replace(/\^\{([^}]+)\}/g, (_: string, p1: string) => charMap(p1, SUPERSCRIPTS))
    // Subscript with braces — handles multi-char like _{n+1}, _{max}
    .replace(/_\{([^}]+)\}/g, (_: string, p1: string) => charMap(p1, SUBSCRIPTS))
    // Clean up remaining braces
    .replace(/[{}]/g, '')
    // Spaces around operators
    .replace(/\\,/g, ' ')
    .replace(/\\;/g, '  ')
    .replace(/\\quad/g, '    ')
    .replace(/\\qquad/g, '        ')
    .trim();
  return result;
}

/**
 * Check if a string looks like LaTeX math.
 */
export function isLatexMath(text: string): boolean {
  return /\\[a-zA-Z]+|\^\{|_\{|\\frac|\\sum|\\int|\\sqrt/.test(text);
}

/**
 * Render a LaTeX equation for DOCX insertion.
 * V2: Returns OMML XML string when pandoc available, Unicode otherwise.
 */
export function renderEquation(latex: string): string {
  return latexToOMML(latex);
}

// ── Native math AST ──────────────────────────────────────

class LatexMathParser {
  private index = 0;

  constructor(private readonly source: string) {}

  parse(stopCharacter?: string): LatexMathNode[] {
    const nodes: LatexMathNode[] = [];
    while (this.index < this.source.length) {
      if (stopCharacter && this.source[this.index] === stopCharacter) break;
      const atom = this.parseAtom();
      if (atom.length === 0) continue;

      let subScript: LatexMathNode[] | undefined;
      let superScript: LatexMathNode[] | undefined;
      this.skipWhitespace();
      while (this.source[this.index] === '_' || this.source[this.index] === '^') {
        const marker = this.source[this.index];
        this.index += 1;
        const value = this.parseScriptValue();
        if (marker === '_') subScript = value;
        else superScript = value;
        this.skipWhitespace();
      }

      if (subScript || superScript) {
        const only = atom.length === 1 ? atom[0] : undefined;
        if (only?.type === 'sum' || only?.type === 'integral') {
          appendMathNode(nodes, { ...only, subScript, superScript });
        } else {
          appendMathNode(nodes, { type: 'script', base: atom, subScript, superScript });
        }
      } else {
        for (const node of atom) appendMathNode(nodes, node);
      }
    }
    return nodes;
  }

  private parseAtom(): LatexMathNode[] {
    const character = this.source[this.index];
    if (character === undefined) return [];
    if (/\s/.test(character)) {
      this.skipWhitespace();
      return [{ type: 'run', text: ' ' }];
    }
    if (character === '{') {
      this.index += 1;
      const children = this.parse('}');
      if (this.source[this.index] === '}') this.index += 1;
      return children;
    }
    if (character !== '\\') {
      this.index += 1;
      return [{ type: 'run', text: character }];
    }

    this.index += 1;
    const command = this.readCommand();
    if (!command) return [];
    if (command === 'frac' || command === 'dfrac' || command === 'tfrac') {
      return [{
        type: 'fraction',
        numerator: this.parseRequiredGroup(),
        denominator: this.parseRequiredGroup(),
      }];
    }
    if (command === 'sqrt') {
      const degree = this.parseOptionalGroup('[', ']');
      return [{ type: 'radical', children: this.parseRequiredGroup(), degree }];
    }
    if (command === 'sum') return [{ type: 'sum' }];
    if (command === 'int' || command === 'iint' || command === 'iiint') {
      const integralCount = command === 'int' ? 1 : command === 'iint' ? 2 : 3;
      return Array.from({ length: integralCount }, () => ({ type: 'integral' as const }));
    }
    if (['text', 'textrm', 'mathrm', 'mathbf', 'mathit', 'operatorname'].includes(command)) {
      return this.parseRequiredGroup();
    }
    if (command === 'left' || command === 'right') return this.parseAtom();
    if (command === ',' || command === ':' || command === ';' || command === 'quad') {
      return [{ type: 'run', text: command === 'quad' ? '    ' : ' ' }];
    }
    if (command === 'qquad') return [{ type: 'run', text: '        ' }];
    if (command === '!' || command === ' ') return [];

    return [{ type: 'run', text: MATH_SYMBOLS[command] ?? command }];
  }

  private parseRequiredGroup(): LatexMathNode[] {
    this.skipWhitespace();
    if (this.source[this.index] !== '{') return this.parseAtom();
    this.index += 1;
    const children = this.parse('}');
    if (this.source[this.index] === '}') this.index += 1;
    return children;
  }

  private parseOptionalGroup(open: string, close: string): LatexMathNode[] | undefined {
    this.skipWhitespace();
    if (this.source[this.index] !== open) return undefined;
    this.index += 1;
    const children = this.parse(close);
    if (this.source[this.index] === close) this.index += 1;
    return children;
  }

  private parseScriptValue(): LatexMathNode[] {
    this.skipWhitespace();
    return this.parseRequiredGroup();
  }

  private readCommand(): string {
    const start = this.index;
    while (/[A-Za-z]/.test(this.source[this.index] ?? '')) this.index += 1;
    if (this.index > start) return this.source.slice(start, this.index);
    const command = this.source[this.index] ?? '';
    this.index += command ? 1 : 0;
    return command;
  }

  private skipWhitespace(): void {
    while (/\s/.test(this.source[this.index] ?? '')) this.index += 1;
  }
}

function appendMathNode(nodes: LatexMathNode[], node: LatexMathNode): void {
  const previous = nodes[nodes.length - 1];
  if (previous?.type === 'run' && node.type === 'run') previous.text += node.text;
  else nodes.push(node);
}

const MATH_SYMBOLS: Record<string, string> = {
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
  Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ',
  tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  infinity: '∞', infty: '∞', partial: '∂', nabla: '∇', prod: '∏',
  propto: '∝', times: '×', cdot: '·', pm: '±', mp: '∓', leq: '≤', geq: '≥',
  neq: '≠', approx: '≈', equiv: '≡', sim: '∼', parallel: '∥', perp: '⊥',
  to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐',
  ldots: '…', cdots: '⋯', ell: 'ℓ', hbar: 'ℏ', Re: 'ℜ', Im: 'ℑ',
};

// ── OMML generation via pandoc ────────────────────────────

function generateOMMLviaPandoc(latex: string): string | null {
  let mdPath: string | undefined;
  let docxPath: string | undefined;
  try {
    // Create a minimal Markdown file with the equation
    const tmpDir = join(tmpdir(), 'openthesis-omml');
    if (!existsSync(tmpDir)) mkdirSync(tmpDir, { recursive: true });

    const id = randomUUID().slice(0, 8);
    mdPath = join(tmpDir, `eq-${id}.md`);
    docxPath = join(tmpDir, `eq-${id}.docx`);

    // Wrap in display math for pandoc
    const safeLatex = latex.replace(/`/g, '\\`');
    writeFileSync(mdPath, `$$${safeLatex}$$`, 'utf-8');

    // Run pandoc
    execFileSync('pandoc', [mdPath, '-o', docxPath], {
      encoding: 'utf-8',
      stdio: 'pipe',
      timeout: 5000,
    });

    // Extract OMML from the generated DOCX
    const omml = extractOMML(docxPath);

    return omml;
  } catch {
    return null;  // pandoc not available, fall back to Unicode
  } finally {
    for (const path of [mdPath, docxPath]) {
      if (path && existsSync(path)) {
        try { unlinkSync(path); } catch {}
      }
    }
  }
}

function extractOMML(docxPath: string): string | null {
  try {
    // DOCX is a ZIP file
    const zip = new AdmZip(docxPath);
    const docXml = zip.readAsText('word/document.xml');

    // Extract the first oMath element
    const match = docXml.match(/<m:oMath[\s\S]*?<\/m:oMath>/);
    return match ? match[0] : null;
  } catch {
    return null;
  }
}

// ── Unicode helpers ───────────────────────────────────────

const SUPERSCRIPTS: Record<string, string> = {
  '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹',
  '+':'⁺','-':'⁻','=':'⁼','(':'⁽',')':'⁾',
  'a':'ᵃ','b':'ᵇ','c':'ᶜ','d':'ᵈ','e':'ᵉ','f':'ᶠ','g':'ᵍ','h':'ʰ',
  'i':'ⁱ','j':'ʲ','k':'ᵏ','l':'ˡ','m':'ᵐ','n':'ⁿ','o':'ᵒ','p':'ᵖ',
  'r':'ʳ','s':'ˢ','t':'ᵗ','u':'ᵘ','v':'ᵛ','w':'ʷ','x':'ˣ','y':'ʸ','z':'ᶻ',
  'A':'ᴬ','B':'ᴮ','D':'ᴰ','E':'ᴱ','G':'ᴳ','H':'ᴴ','I':'ᴵ','J':'ᴶ',
  'K':'ᴷ','L':'ᴸ','M':'ᴹ','N':'ᴺ','O':'ᴼ','P':'ᴾ','R':'ᴿ','T':'ᵀ',
  'U':'ᵁ','V':'ⱽ','W':'ᵂ',
  'α':'ᵅ','β':'ᵝ','γ':'ᵞ','δ':'ᵟ','ε':'ᵋ','θ':'ᶿ',
  'φ':'ᵠ','χ':'ᵡ',
};

const SUBSCRIPTS: Record<string, string> = {
  '0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉',
  '+':'₊','-':'₋','=':'₌','(':'₍',')':'₎',
  'a':'ₐ','e':'ₑ','h':'ₕ','i':'ᵢ','j':'ⱼ','k':'ₖ','l':'ₗ',
  'm':'ₘ','n':'ₙ','o':'ₒ','p':'ₚ','r':'ᵣ','s':'ₛ','t':'ₜ',
  'u':'ᵤ','v':'ᵥ','x':'ₓ',
  // Greek subscripts (limited Unicode support)
  'β':'ᵦ','γ':'ᵧ','ρ':'ᵨ','φ':'ᵩ','χ':'ᵪ',
};

function charMap(text: string, map: Record<string, string>): string {
  return [...text].map(c => map[c] || c).join('');
}
