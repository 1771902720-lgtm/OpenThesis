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
