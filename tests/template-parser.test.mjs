import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { parseTemplate } from '../packages/template-parser/dist/index.js';

async function createTemplate() {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('parses page settings and resolves inherited styles', async () => {
  const template = await parseTemplate(await createTemplate(), {
    organization: 'Test University', documentType: 'thesis',
  });
  assert.equal(template.meta.organization, 'Test University');
  assert.equal(template.page.columns, 2);
  assert.equal(template.page.columnGutter, 720);
  assert.equal(template.styles.Heading1.font.name, 'Arial');
  assert.equal(template.styles.Heading1.font.size, 32);
  assert.equal(template.styles.Heading1.font.bold, true);
  assert.equal(template.styleRoles.Heading1, 'heading1');
  assert.equal(template.styleInheritance.Heading1, 'Normal');
});

test('rejects a ZIP that is not a Word template', async () => {
  const zip = new JSZip();
  zip.file('readme.txt', 'not a docx');
  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  await assert.rejects(() => parseTemplate(buffer), /No styles\.xml found/);
});

test('keeps metadata defaults when optional overrides are undefined', async () => {
  const template = await parseTemplate(await createTemplate(), {
    organization: undefined,
    name: undefined,
  });
  assert.equal(template.meta.organization, 'Unknown Organization');
  assert.equal(template.meta.name, 'Untitled Template');
  assert.equal(template.meta.parserVersion, '0.2.0');
});

/**
 * Build a template whose heading carries its font in `w:pPr/w:rPr` (the shape
 * Chinese Word/WPS templates emit) and whose indents use the `*Chars` units.
 */
async function createChineseTemplate() {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault/></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/><w:ind w:firstLineChars="200"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/>
        <w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="32"/></w:rPr></w:pPr>
      </w:style>
      <w:style w:type="character" w:styleId="uChar"><w:name w:val="char"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>
      <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
      <w:style w:type="paragraph" w:styleId="MTConvertedEquation"><w:name w:val="MT Converted Equation"/><w:basedOn w:val="a"/><w:rPr><w:i/><w:sz w:val="28"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1797" w:right="1797"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('reads run properties nested under w:pPr/w:rPr', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // Without this, every style looked empty and headings rendered as body text.
  assert.equal(template.styles['1'].font.eastAsia, '黑体');
  assert.equal(template.styles['1'].font.size, 32);
  assert.equal(template.styles['1'].font.bold, true);
  assert.equal(template.styles['1'].paragraph.alignment, 'center');
});

test('converts w:firstLineChars into twips instead of storing the raw count', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // 200 hundredths of a character at 12pt = 2 chars × 240 twips = 480 twips.
  assert.equal(template.styles.a.paragraph.firstLineIndent, 480);
});

test('only paragraph styles receive block roles', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // Character and table styles are not paragraph-level blocks; assigning them a
  // role let `resolveStyle` pick e.g. a character style for body text.
  assert.equal(template.styleRoles.uChar, undefined);
  assert.equal(template.styleRoles.TableGrid, undefined);
  // `Table Grid` used to become heading3 via an unanchored /^(表|Table|题注)/.
  assert.notEqual(template.styleRoles.TableGrid, 'heading3');
});

test('detects a template equation style', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  assert.equal(template.styleRoles.MTConvertedEquation, 'equation');
});

test('omits properties a style does not declare', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // `Normal` declares no run properties, so none may be invented for it —
  // otherwise "the template is silent" is indistinguishable from an explicit
  // choice, and the renderer's role defaults can never take effect.
  assert.equal(template.styles.a.font.size, undefined);
  assert.equal(template.styles.a.font.eastAsia, undefined);
  assert.equal(template.styles.a.font.name, undefined);
  // Properties it does declare are kept.
  assert.equal(template.styles.a.paragraph.alignment, 'justified');
  // Nothing is invented for absent spacing either.
  assert.equal(template.styles['1'].paragraph.spaceBefore, undefined);
});
