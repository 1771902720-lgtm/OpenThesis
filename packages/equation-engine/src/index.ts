// ============================================================
// @openthesis/equation-engine — LaTeX → Office Math Engine
// ============================================================
// V2: Unicode fallback + pandoc OMML integration
// When pandoc is available, generates native Word OMML equations.
// Falls back to Unicode plain-text conversion otherwise.
// ============================================================

import { execSync } from 'child_process';
import { writeFileSync, unlinkSync, existsSync, mkdirSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { randomUUID } from 'crypto';

// ── Public API ────────────────────────────────────────────

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

// ── OMML generation via pandoc ────────────────────────────

function generateOMMLviaPandoc(latex: string): string | null {
  try {
    // Create a minimal Markdown file with the equation
    const tmpDir = join(tmpdir(), 'openthesis-omml');
    if (!existsSync(tmpDir)) mkdirSync(tmpDir, { recursive: true });

    const id = randomUUID().slice(0, 8);
    const mdPath = join(tmpDir, `eq-${id}.md`);
    const docxPath = join(tmpDir, `eq-${id}.docx`);

    // Wrap in display math for pandoc
    const safeLatex = latex.replace(/`/g, '\\`');
    writeFileSync(mdPath, `$$${safeLatex}$$`, 'utf-8');

    // Run pandoc
    execSync(`pandoc "${mdPath}" -o "${docxPath}"`, {
      encoding: 'utf-8',
      stdio: 'pipe',
      timeout: 5000,
    });

    // Extract OMML from the generated DOCX
    const omml = extractOMML(docxPath);

    // Cleanup
    try { unlinkSync(mdPath); unlinkSync(docxPath); } catch {}

    return omml;
  } catch {
    return null;  // pandoc not available, fall back to Unicode
  }
}

function extractOMML(docxPath: string): string | null {
  try {
    // DOCX is a ZIP file
    const AdmZip = require('adm-zip');
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
