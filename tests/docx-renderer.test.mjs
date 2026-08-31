import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createUSTBTemplate, renderDocument } from '../packages/docx-renderer/dist/index.js';

const template = createUSTBTemplate();

async function documentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return { zip, xml: await zip.file('word/document.xml').async('string') };
}

test('renders real list, code, quote and figure blocks with a caption', async () => {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-renderer-'));
  writeFileSync(
    join(contentDir, 'figure.png'),
    Buffer.from('/9j/4AAQSkZJRgABAQAAAAAAAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/wAALCAACAAIBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVN//2Q==', 'base64'),
  );
  const buffer = await renderDocument({
    outputPath: '',
    template,
    contentDir,
    document: {
      type: 'thesis',
      meta: { title: 'Renderer test' },
      cover: [],
      sections: [{
        id: 'blocks', type: 'chapter', title: 'Blocks',
        content: [
          { type: 'list_item', text: 'First item' },
          { type: 'code_block', text: 'const x = 1;\nconsole.log(x);' },
          { type: 'blockquote', text: 'Quoted text' },
          { type: 'figure', path: 'figure.png', caption: 'Figure caption', number: '1' },
        ],
      }],
    },
  });
  const { zip, xml } = await documentXml(buffer);
  assert.match(xml, /First item/);
  assert.match(xml, /const x = 1;/);
  assert.match(xml, /Quoted text/);
  assert.match(xml, /1  Figure caption/);
  assert.doesNotMatch(xml, /\[list_item:|\[code_block:|\[blockquote:/);
  const relationships = await zip.file('word/_rels/document.xml.rels').async('string');
  if (/relationships\/theme/.test(relationships)) {
    assert.ok(zip.file('word/theme/theme1.xml'), 'a theme relationship must keep its target part');
  }

  const mediaName = Object.keys(zip.files).find(name => name.startsWith('word/media/') && !name.endsWith('/'));
  assert.ok(mediaName, 'figure image should be embedded');
  const media = await zip.file(mediaName).async('nodebuffer');
  assert.equal(media.readUInt16BE(0), 0xffd8, 'JPEG bytes should be detected despite the .png filename');
});

test('renders the CC label exactly once in official documents', async () => {
  const buffer = await renderDocument({
    outputPath: '',
    template,
    document: {
      type: 'official',
      meta: {
        title: '测试通知', issuingAuthority: '测试机关', documentNumber: '测〔2026〕1号',
        documentCategory: '通知', primaryRecipients: ['主送单位'], ccRecipients: ['抄送单位'], date: '2026年8月30日',
      },
      body: [{ type: 'paragraph', text: '正文。' }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /抄送：抄送单位。/);
  assert.doesNotMatch(xml, /抄送：抄送：/);
});

test('rejects tables without header columns', async () => {
  await assert.rejects(() => renderDocument({
    outputPath: '',
    template,
    document: {
      type: 'thesis', meta: { title: 'Bad table' }, cover: [],
      sections: [{
        id: 'table', type: 'chapter', title: 'Table',
        content: [{ type: 'table', caption: 'Empty', headers: [], data: [] }],
      }],
    },
  }), /must define at least one header column/);
});

test('rejects table column width counts that do not match headers', async () => {
  await assert.rejects(() => renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Bad widths' }, cover: [],
      sections: [{
        id: 'table', type: 'chapter', title: 'Table',
        content: [{ type: 'table', caption: 'Widths', headers: ['A', 'B'], data: [], columnWidths: [1000] }],
      }],
    },
  }), /2 columns but 1 column widths/);
});

test('writes configured multi-column sections to OOXML', async () => {
  const twoColumnTemplate = structuredClone(template);
  twoColumnTemplate.page.columns = 2;
  twoColumnTemplate.page.columnGutter = 720;
  const buffer = await renderDocument({
    outputPath: '', template: twoColumnTemplate,
    document: {
      type: 'thesis', meta: { title: 'Columns' }, cover: [],
      sections: [{ id: 'body', type: 'chapter', title: 'Body', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:cols[^>]*w:space="720"[^>]*w:num="2"|<w:cols[^>]*w:num="2"[^>]*w:space="720"/);
});

test('writes supported LaTeX as native Office Math elements', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Native math' }, cover: [],
      sections: [{
        id: 'math', type: 'chapter', title: 'Math',
        content: [{ type: 'equation_numbered', latex: '\\frac{x_1}{\\sqrt{y^2}}', number: '1' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<m:oMath>/);
  assert.match(xml, /<m:f>/);
  assert.match(xml, /<m:rad>/);
  assert.match(xml, /<m:sSub>/);
  assert.match(xml, /<m:sSup>/);
  assert.match(xml, /\(1\)/);
});

test('writes advanced LaTeX macros and environments as native OMML', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Advanced native math' }, cover: [],
      sections: [{
        id: 'math', type: 'chapter', title: 'Math',
        content: [
          { type: 'equation', latex: '\\hat{x}+\\lim_{t\\to0}\\sin(t)+\\binom{n}{k}' },
          { type: 'equation', latex: '\\overline{AB}+\\prod_{i=1}^{n}x_i+\\overset{*}{=}+\\left\\langle x,y\\right\\rangle' },
          { type: 'equation', latex: '\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}' },
          { type: 'equation', latex: 'f(x)=\\begin{cases}x^2 & x\\geq0 \\\\ -x & x<0\\end{cases}' },
        ],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<m:acc>/);
  assert.match(xml, /<m:bar>/);
  assert.match(xml, /<m:limLow>/);
  assert.match(xml, /<m:limUpp>/);
  assert.match(xml, /<m:func>/);
  assert.match(xml, /<m:nary>/);
  assert.match(xml, /<m:chr m:val="∏"\/>/);
  assert.match(xml, /<m:type m:val="noBar"\/>/);
  assert.match(xml, /<m:m>/);
  assert.match(xml, /<m:mr>/);
  assert.match(xml, /<m:begChr m:val="\("\/>/);
  assert.match(xml, /<m:begChr m:val="\{"\/>/);
  assert.match(xml, /<m:begChr m:val="⟨"\/>/);
  assert.match(xml, /<m:endChr m:val=""\/>/);
});
