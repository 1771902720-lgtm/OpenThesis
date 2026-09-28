import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
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

test('carries an exact line-spacing rule into the OOXML', async () => {
  const exactTemplate = structuredClone(template);
  // A role is resolved through `roleWinners`; "the first style carrying the
  // role" is a different style, so mutating that one changed nothing.
  const paragraphStyleId = exactTemplate.roleWinners?.paragraph
    ?? Object.entries(exactTemplate.styleRoles).find(([, role]) => role === 'paragraph')[0];
  exactTemplate.styles[paragraphStyleId] = {
    ...exactTemplate.styles[paragraphStyleId],
    lineSpacing: 360,
    lineSpacingRule: 'exact',
  };

  const buffer = await renderDocument({
    outputPath: '', template: exactTemplate,
    document: {
      type: 'thesis', meta: { title: 'Spacing' }, cover: [],
      sections: [{
        id: 'body', type: 'chapter', title: 'Body',
        content: [{ type: 'paragraph', text: 'Body text.' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Without the rule Word reads 360 as 1.5 lines instead of 18 points.
  assert.match(
    xml,
    /<w:spacing[^>]*w:line="360"[^>]*w:lineRule="exact"|<w:spacing[^>]*w:lineRule="exact"[^>]*w:line="360"/,
  );
});

test('finds a style role without scanning the whole table per block', async () => {
  // `Object.entries(styleRoles).find(...)` materialised all 72 pairs before
  // matching, which cost 9-11µs per block — about a fifth of render time.
  let reads = 0;
  const countingTemplate = structuredClone(template);
  countingTemplate.styleRoles = new Proxy(countingTemplate.styleRoles, {
    get(target, key, receiver) {
      reads += 1;
      return Reflect.get(target, key, receiver);
    },
  });

  const blocks = Array.from({ length: 40 }, (_, index) => ({ type: 'paragraph', text: `Paragraph ${index}.` }));
  await renderDocument({
    outputPath: '', template: countingTemplate,
    document: {
      type: 'thesis', meta: { title: 'Role lookup' }, cover: [],
      sections: [{ id: 'body', type: 'chapter', title: 'Body', content: blocks }],
    },
  });

  // A per-block scan reads every role for every block (≈2900 here); stopping at
  // the first match reads a handful. Counting property reads keeps the test
  // independent of machine speed.
  assert.ok(reads < blocks.length * 5, `${reads} role reads for ${blocks.length} blocks — the table is scanned per block`);
});

test('the built-in USTB template carries the university formatting', () => {
  // The asset was generated by a parser that invented one signature for all 72
  // styles, so every heading rendered like body text. These are the numbers the
  // 《北京科技大学研究生学位论文书写指南》 asks for.
  const asset = JSON.parse(readFileSync('assets/ustb-thesis-template.json', 'utf8'));
  const style = role => asset.styles[asset.roleWinners[role]];

  assert.equal(style('heading1').font.eastAsia, '黑体');
  assert.equal(style('heading1').font.size, 30);          // 小三
  assert.equal(style('heading1').paragraph.alignment, 'center');
  assert.equal(style('heading2').font.size, 28);          // 四号
  assert.equal(style('heading3').font.size, 28);
  assert.equal(style('paragraph').font.eastAsia, '宋体');
  assert.equal(style('paragraph').font.size, 24);         // 小四
  assert.equal(style('paragraph').paragraph.firstLineIndent, 480);   // 2 characters
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

test('emits real Word list numbering instead of literal markers', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Lists' }, cover: [],
      sections: [{
        id: 'l', type: 'chapter', title: 'Lists',
        content: [
          { type: 'list_item', text: 'bullet one' },
          { type: 'list_item', text: 'bullet two', level: 1 },
          { type: 'paragraph', text: 'interrupt' },
          { type: 'list_item', text: 'second bullet list' },
          { type: 'paragraph', text: 'interrupt again' },
          { type: 'list_item', text: 'ordered one', ordered: true },
          { type: 'list_item', text: 'ordered two', ordered: true },
        ],
      }],
    },
  });

  const { zip, xml } = await documentXml(buffer);
  const numbering = await zip.file('word/numbering.xml')?.async('string');

  assert.ok(numbering, 'numbering.xml must be written');
  assert.match(numbering, /<w:abstractNum/);
  assert.match(numbering, /w:val="•"/);

  const numPrs = [...xml.matchAll(/<w:numPr>[\s\S]*?<\/w:numPr>/g)];
  assert.equal(numPrs.length, 5);

  // Each separate list restarts its own counter, so the two bullet lists must
  // not share a numId.
  const numIds = [...xml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map(m => m[1]);
  assert.ok(new Set(numIds).size >= 3, `expected distinct numIds, got ${JSON.stringify(numIds)}`);

  const levels = [...xml.matchAll(/<w:ilvl w:val="(\d+)"\/>/g)].map(m => m[1]);
  assert.ok(levels.includes('1'), 'nested level must be written');

  // The old renderer prefixed the text with a literal marker.
  assert.doesNotMatch(xml, /<w:t[^>]*>•\s/);
  assert.match(xml, /bullet one/);
  assert.match(xml, /ordered two/);
});

test('applies the guide formatting the built-in template carries', async () => {
  // One document that exercises the properties added for the university guide,
  // asserted on the produced OOXML rather than on the template object.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Guide check' }, cover: [],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [{ type: 'paragraph', text: '摘要正文。' }] },
        { id: 'ch', type: 'chapter', title: '第一章', content: [{ type: 'paragraph', text: '正文。' }] },
      ],
      backMatter: { references: [{ id: '1', text: '一条参考文献.' }] },
    },
  });
  const { xml } = await documentXml(buffer);

  // 对称页边距 + 装订线 1cm (567 twips) for a bound thesis.
  assert.match(xml, /<w:mirrorMargins\/>/);
  assert.match(xml, /w:gutter="567"/);
  // A section heading is not a chapter heading: 段前17磅/段后16.5磅/2.41倍行距(579),
  // where heading1 is 340/340 and 1.3倍(312).
  assert.match(xml, /w:line="579"/);
  // A hanging indent belongs to a flush-left heading (level 2 = 1cm) and to a
  // reference entry (1cm, measured in the real thesis). It must NOT sit on a
  // centred heading: the guide's "0.75cm" cannot survive contact with a centred
  // line, and the accepted thesis has every chapter heading exactly centred.
  assert.match(xml, /w:hanging="567"/);
  assert.doesNotMatch(xml, /w:hanging="425"/);
  // Body: 宋体 小四 with a two-character first-line indent at 1.3 line spacing.
  assert.match(xml, /w:firstLine="480"/);
  assert.match(xml, /w:line="312"/);
});

test('styles the cover from the template, one tier at a time', async () => {
  // The guide specifies the cover box by box: 校名行 小二 18pt bold centred,
  // 研究生/指导教师 四号 14pt bold on justified rows, 中图分类号 五号. The renderer
  // used to hardcode centre and one size for every cover line.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Cover' },
      cover: [
        { type: 'centered_text', text: '北京科技大学博士学位论文', styleRole: 'cover_title' },
        { type: 'centered_text', text: '研究生  张三', styleRole: 'cover_line' },
        { type: 'centered_text', text: '中图分类号：TU3', styleRole: 'cover_meta' },
      ],
      sections: [{ id: 'a', type: 'chapter', title: '第一章', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);

  const para = text => xml.split(/<\/w:p>/).find(chunk => chunk.includes(text)) ?? '';
  const title = para('北京科技大学博士学位论文');
  assert.match(title, /<w:sz w:val="36"\/>/);       // 小二
  assert.match(title, /<w:b\/>/);
  assert.match(title, /<w:jc w:val="center"\/>/);

  const line = para('研究生  张三');
  assert.match(line, /<w:sz w:val="28"\/>/);        // 四号
  assert.match(line, /<w:jc w:val="both"\/>/);      // the guide's 两端对齐 row

  const meta = para('中图分类号：TU3');
  assert.match(meta, /<w:sz w:val="21"\/>/);        // 五号
  assert.doesNotMatch(meta, /<w:b\/>/);
});

test('emits a real TOC field for a 目录 section', async () => {
  // A table of contents is a field, not text: Word fills in the entries and
  // their page numbers when the field is updated. No generator can compute page
  // numbers, so it emits the field rather than faking a list.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'TOC' }, cover: [],
      sections: [
        { id: 'toc', type: 'toc', title: '目录', content: [] },
        { id: 'c1', type: 'chapter', title: '第一章 绪论', content: [{ type: 'paragraph', text: '正文。' }] },
      ],
    },
  });
  const { xml } = await documentXml(buffer);
  // The field's quotes are XML-escaped in the part.
  assert.match(xml, /<w:instrText[^>]*>TOC \\h \\o &quot;1-3&quot;<\/w:instrText>/);
  assert.match(xml, /目录/);
  assert.match(xml, /第一章 绪论/);
});
