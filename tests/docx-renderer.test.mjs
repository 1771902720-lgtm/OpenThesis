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

/** The `<w:rPr>` of the run containing `needle`. */
function runProperties(xml, needle) {
  const index = xml.indexOf(needle);
  assert.notEqual(index, -1, `expected to find ${needle}`);
  const start = xml.lastIndexOf('<w:r>', index);
  return xml.slice(start, xml.indexOf('</w:r>', index) + 6);
}

test('renders a fourth-level heading as a heading, not as body text', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Nesting' }, cover: [],
      sections: [{
        id: 'l1', type: 'chapter', title: 'Level One', content: [],
        subsections: [{
          id: 'l2', type: 'section', title: 'Level Two', content: [],
          subsections: [{
            id: 'l3', type: 'section', title: 'Level Three', content: [],
            subsections: [{ id: 'l4', type: 'section', title: 'Level Four', content: [] }],
          }],
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // `heading4` was missing from the fallback table, so it resolved to the body
  // paragraph style and looked identical to surrounding text.
  const body = runProperties(xml, 'Level One');
  const fourth = runProperties(xml, 'Level Four');
  assert.match(fourth, /<w:b\/>/);
  assert.doesNotMatch(fourth, /w:eastAsia="仿宋"/);
  assert.match(fourth, /w:eastAsia="黑体"/);
  assert.notEqual(fourth, body);
});

test('fits a table to the printable page width instead of a fixed A4 constant', async () => {
  const narrow = structuredClone(template);
  narrow.page.width = 8391;
  narrow.page.height = 11906;
  narrow.page.margins = { top: 1701, bottom: 1701, left: 1701, right: 1701 };
  const printable = narrow.page.width - narrow.page.margins.left - narrow.page.margins.right;

  const buffer = await renderDocument({
    outputPath: '', template: narrow,
    document: {
      type: 'thesis', meta: { title: 'Table width' }, cover: [],
      sections: [{
        id: 't', type: 'chapter', title: 'T',
        content: [{ type: 'table', caption: 'Cap', headers: ['A', 'B'], data: [['1', '2']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Cell widths always summed to 8504 twips, overflowing a narrower page.
  const cellWidths = [...xml.matchAll(/<w:tcW w:type="dxa" w:w="(\d+)"/g)].map(m => Number(m[1]));
  const rowWidth = cellWidths.slice(0, 2).reduce((sum, width) => sum + width, 0);
  assert.ok(rowWidth <= printable, `row width ${rowWidth} must fit printable ${printable}`);
  assert.match(xml, new RegExp(`<w:tblW w:type="dxa" w:w="${rowWidth}"`));
  assert.doesNotMatch(xml, /<w:tblW w:type="pct"/);
});

test('takes table typography from the template', async () => {
  const styled = structuredClone(template);
  styled.styles = {
    ...styled.styles,
    MyTableText: { font: { name: 'Georgia', eastAsia: '楷体', size: 36 }, paragraph: { alignment: 'left' } },
    MyTableHead: { font: { name: 'Georgia', eastAsia: '黑体', size: 36, bold: true }, paragraph: { alignment: 'left' } },
  };
  styled.styleRoles = { ...styled.styleRoles, MyTableText: 'table', MyTableHead: 'table_header' };

  const buffer = await renderDocument({
    outputPath: '', template: styled,
    document: {
      type: 'thesis', meta: { title: 'Table style' }, cover: [],
      sections: [{
        id: 't', type: 'chapter', title: 'T',
        content: [{ type: 'table', caption: 'Cap', headers: ['HDR'], data: [['DAT']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Tables used to hardcode Times New Roman / 黑体 / 宋体 and ignore the template.
  assert.match(runProperties(xml, 'HDR'), /w:ascii="Georgia"/);
  assert.match(runProperties(xml, 'HDR'), /w:eastAsia="黑体"/);
  assert.match(runProperties(xml, 'HDR'), /<w:sz w:val="36"\/>/);
  assert.match(runProperties(xml, 'DAT'), /w:eastAsia="楷体"/);
  assert.match(runProperties(xml, 'DAT'), /<w:sz w:val="36"\/>/);
});

test('falls back to domain-specific typography when the template is silent', async () => {
  const bare = structuredClone(template);
  bare.styles = {};
  bare.styleRoles = {};

  const thesisBuffer = await renderDocument({
    outputPath: '', template: bare,
    document: {
      type: 'thesis', meta: { title: 'Defaults' }, cover: [],
      sections: [{ id: 'p', type: 'chapter', title: 'Chapter', content: [{ type: 'paragraph', text: 'BodyText' }] }],
    },
  });
  const officialBuffer = await renderDocument({
    outputPath: '', template: bare,
    document: {
      type: 'official',
      meta: {
        title: '公文', issuingAuthority: '某机关', documentNumber: 'X〔2026〕1号',
        documentCategory: '通知', primaryRecipients: ['各单位'], date: '2026年1月1日',
      },
      body: [{ type: 'paragraph', text: 'BodyText' }],
    },
  });

  const thesisXml = (await documentXml(thesisBuffer)).xml;
  const officialXml = (await documentXml(officialBuffer)).xml;

  // 学位论文 body is 宋体 小四; GB/T 9704 公文 body is 仿宋 三号.
  assert.match(runProperties(thesisXml, 'BodyText'), /w:eastAsia="宋体"/);
  assert.match(runProperties(thesisXml, 'BodyText'), /<w:sz w:val="24"\/>/);
  assert.match(runProperties(officialXml, 'BodyText'), /w:eastAsia="仿宋"/);
  assert.match(runProperties(officialXml, 'BodyText'), /<w:sz w:val="32"\/>/);
  assert.match(runProperties(thesisXml, 'Chapter'), /w:eastAsia="黑体"/);
});
