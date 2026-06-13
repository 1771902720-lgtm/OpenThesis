#!/usr/bin/env node
// ============================================================
// @openthesis/cli — OpenThesis Command Line Interface
// ============================================================
// Usage:
//   thesis parse <template.docx> [--type thesis|journal|official]
//   thesis build <content.json> [-t <template.json>] [-o <output.docx>]
//   thesis init [--type thesis|journal|official]
// ============================================================

import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';
import { parseTemplate } from '@openthesis/template-parser';
import { renderLegacy, createUSTBTemplate } from '@openthesis/docx-renderer';
import type { DocumentTemplate, DocumentType, LegacyDocumentJSON } from '@openthesis/document-schema';

const args = process.argv.slice(2);
const command = args[0];

async function main() {
  switch (command) {
    case 'parse':   await cmdParse(args.slice(1)); break;
    case 'build':   await cmdBuild(args.slice(1)); break;
    case 'init':    cmdInit(args.slice(1)); break;
    default:        printHelp(); process.exit(1);
  }
}

// ── thesis parse <template.docx> ──────────────────────────

async function cmdParse(args: string[]) {
  if (args.length < 1) {
    console.error('Usage: thesis parse <template.docx> [--type thesis|journal|official] [--org <name>]');
    process.exit(1);
  }

  const templatePath = resolve(args[0]);
  const typeIdx = args.indexOf('--type');
  const orgIdx = args.indexOf('--org');
  const docType: DocumentType = typeIdx >= 0 ? (args[typeIdx + 1] as DocumentType) : 'thesis';
  const orgName = orgIdx >= 0 ? args[orgIdx + 1] : undefined;

  console.log(`Parsing template: ${templatePath}`);
  console.log(`Document type:   ${docType}`);

  try {
    const buffer = readFileSync(templatePath);
    const template = await parseTemplate(buffer, {
      organization: orgName,
      documentType: docType,
      sourceFile: templatePath,
    });

    const outPath = templatePath.replace(/\.docx?$/i, '.template.json');
    writeFileSync(outPath, JSON.stringify(template, null, 2), 'utf-8');

    console.log(`\n✅ Template parsed successfully!`);
    console.log(`   Output:    ${outPath}`);
    console.log(`   Org:       ${template.meta.organization}`);
    console.log(`   Type:      ${template.meta.documentType}`);
    console.log(`   Page:      ${Math.round(template.page.width / 20)}pt × ${Math.round(template.page.height / 20)}pt`);
    if (template.page.columns && template.page.columns > 1) {
      console.log(`   Columns:   ${template.page.columns}`);
    }
    console.log(`   Styles:    ${Object.keys(template.styles).length} found, ${Object.keys(template.styleRoles).length} roles detected`);
    console.log(`\n   Style roles:`);
    for (const [styleId, role] of Object.entries(template.styleRoles)) {
      const style = template.styles[styleId];
      const fontInfo = style
        ? `${style.font.eastAsia || style.font.name} ${style.font.size / 2}pt`
        : '—';
      console.log(`     ${styleId} → ${role} (${fontInfo})`);
    }
  } catch (err: any) {
    console.error(`\n❌ Failed to parse template: ${err.message}`);
    process.exit(1);
  }
}

// ── thesis build <content.json> ───────────────────────────

async function cmdBuild(args: string[]) {
  if (args.length < 1) {
    console.error('Usage: thesis build <content.json> [-t <template.json>] [-o <output.docx>]');
    process.exit(1);
  }

  const contentPath = resolve(args[0]);
  const tIdx = args.indexOf('-t');
  const oIdx = args.indexOf('-o');

  const templatePath = tIdx >= 0 ? resolve(args[tIdx + 1]) : null;
  const outputPath = oIdx >= 0 ? resolve(args[oIdx + 1]) : contentPath.replace(/\.json$/i, '.docx');

  console.log(`Content:  ${contentPath}`);
  console.log(`Template: ${templatePath || '(built-in USTB)'}`);
  console.log(`Output:   ${outputPath}`);

  try {
    const contentJson = JSON.parse(readFileSync(contentPath, 'utf-8')) as LegacyDocumentJSON;

    let template: DocumentTemplate;
    if (templatePath) {
      template = JSON.parse(readFileSync(templatePath, 'utf-8'));
    } else {
      console.log('  Using built-in USTB template...');
      template = createUSTBTemplate();
    }

    await renderLegacy(contentJson, template, outputPath);

    const docType = contentJson.title?.includes('期刊') ? 'journal article'
      : contentJson.title?.includes('公文') || contentJson.title?.includes('通知') ? 'official document'
      : 'document';

    console.log(`\n✅ ${docType} built successfully!`);
    console.log(`   Output: ${outputPath}`);
  } catch (err: any) {
    console.error(`\n❌ Failed to build: ${err.message}`);
    process.exit(1);
  }
}

// ── thesis init [--type thesis|journal|official] ───────────

function cmdInit(args: string[]) {
  const typeIdx = args.indexOf('--type');
  const docType: DocumentType = typeIdx >= 0 ? (args[typeIdx + 1] as DocumentType) : 'thesis';

  switch (docType) {
    case 'thesis':   createThesisSample(); break;
    case 'journal':  createJournalSample(); break;
    case 'official': createOfficialSample(); break;
    default:         createThesisSample(); break;
  }
}

function createThesisSample() {
  const content: LegacyDocumentJSON = {
    title: '示例学位论文标题',
    report_header: 'XX大学博士/硕士学位论文',
    cover_blocks: [
      { type: 'spacer', lines: 6 },
      { type: 'centered_text', text: 'XX大学', font_size_pt: 26, bold: true },
      { type: 'centered_text', text: 'University Name (English)', font_size_pt: 16, bold: false },
      { type: 'spacer', lines: 2 },
      { type: 'centered_text', text: '示例博士学位论文', font_size_pt: 22, bold: true },
      { type: 'spacer', lines: 4 },
      { type: 'centered_text', text: '学科专业：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '研究方向：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '日    期：2026年6月', font_size_pt: 14, bold: false },
      { type: 'page_break' },
    ],
    body_blocks: [
      { type: 'heading1', text: '第一章  绪论' },
      { type: 'heading2', text: '1.1  研究背景' },
      { type: 'paragraph', text: '正文示例——宋体小四号字，1.5倍行距，首行缩进2字符。西文使用Times New Roman。' },
      { type: 'heading2', text: '1.2  研究目的与意义' },
      { type: 'paragraph', text: '在此撰写研究目的...' },
      { type: 'heading1', text: '第二章  理论基础' },
      { type: 'heading2', text: '2.1  基本方程' },
      { type: 'paragraph', text: '控制方程为：' },
      { type: 'equation', latex: 'E = mc^2' },
      { type: 'heading1', text: '第三章  结论' },
      { type: 'paragraph', text: '研究成果总结...' },
    ],
  };

  writeFileSync('thesis-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created thesis-content.json (学位论文示例)');
  printInitNext('thesis');
}

function createJournalSample() {
  const content: LegacyDocumentJSON = {
    title: '示例期刊论文',
    report_header: 'Journal of Example Research',
    cover_blocks: [],
    body_blocks: [
      { type: 'centered_text', text: '示例期刊论文标题', font_size_pt: 16, bold: true },
      { type: 'spacer', lines: 1 },
      { type: 'centered_text', text: '作者一¹  作者二²  通讯作者一¹*', font_size_pt: 11, bold: false },
      { type: 'spacer', lines: 1 },
      { type: 'centered_text', text: '(1. XX大学XX学院, 北京 100083;  2. XX研究所, 上海 200000)', font_size_pt: 10, bold: false },
      { type: 'spacer', lines: 2 },
      { type: 'heading1', text: '摘  要' },
      { type: 'paragraph', text: '本文研究了XXX问题，采用XXX方法，得出XXX结论。' },
      { type: 'paragraph_no_indent', text: '关键词：关键词一；关键词二；关键词三' },
      { type: 'heading1', text: '1  Introduction' },
      { type: 'paragraph', text: 'Introduction content here. This is a sample journal article body text in the standard IMRaD format.' },
      { type: 'heading2', text: '1.1  Background' },
      { type: 'paragraph', text: 'Background and literature review content...' },
      { type: 'heading1', text: '2  Methods' },
      { type: 'paragraph', text: 'Methodology description...' },
      { type: 'heading1', text: '3  Results and Discussion' },
      { type: 'paragraph', text: 'Results analysis...' },
      { type: 'table', caption: 'Table 1  Experimental results', headers: ['Parameter', 'Value', 'Error'], data: [['p1', '0.123', '±0.001'], ['p2', '0.456', '±0.002']] },
      { type: 'heading1', text: '4  Conclusion' },
      { type: 'paragraph', text: 'Concluding remarks...' },
      { type: 'heading1', text: 'References' },
      { type: 'paragraph_no_indent', text: '[1] Author A, Author B. Title of paper. Journal Name, 2026, 100(1): 1-10.' },
      { type: 'paragraph_no_indent', text: '[2] Author C. Title of book. Publisher, 2025.' },
    ],
  };

  writeFileSync('journal-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created journal-content.json (期刊论文示例)');
  printInitNext('journal');
}

function createOfficialSample() {
  const content: LegacyDocumentJSON = {
    title: 'XX省人民政府关于做好2026年防汛工作的通知',
    report_header: '',
    cover_blocks: [],
    body_blocks: [
      { type: 'red_header', text: 'XX省人民政府文件', font_size_pt: 22 },
      { type: 'document_number', text: 'X政发〔2026〕1号' },
      { type: 'spacer', lines: 2 },
      { type: 'centered_text', text: 'XX省人民政府', font_size_pt: 18, bold: true },
      { type: 'centered_text', text: '关于做好2026年防汛工作的通知', font_size_pt: 18, bold: true },
      { type: 'spacer', lines: 1 },
      { type: 'recipient_line', text: '各市、州人民政府，省政府各部门：', recipientType: 'primary' },
      { type: 'paragraph', text: '为切实做好2026年防汛工作，保障人民群众生命财产安全，现将有关事项通知如下：' },
      { type: 'heading1', text: '一、提高思想认识，压实防汛责任' },
      { type: 'paragraph', text: '各级各部门要充分认识当前防汛形势的严峻性，坚决克服麻痹思想和侥幸心理。要严格落实以行政首长负责制为核心的各项防汛责任制，层层传导压力，确保责任到人、措施到位。' },
      { type: 'heading1', text: '二、加强监测预警，做好应急准备' },
      { type: 'paragraph', text: '气象、水文部门要加强监测预报，及时发布预警信息。各地要修订完善防汛应急预案，充实抢险救援队伍，储备充足的防汛物资。' },
      { type: 'heading1', text: '三、突出重点区域，排查风险隐患' },
      { type: 'paragraph', text: '要对水库、山塘、河道堤防、山洪地质灾害易发区、城市低洼地带等重点区域开展拉网式排查，建立隐患台账，限期整改销号。' },
      { type: 'heading1', text: '四、加强值班值守，确保信息畅通' },
      { type: 'paragraph', text: '严格执行24小时值班和领导带班制度，及时准确上报汛情、险情、灾情信息，不得迟报、漏报、瞒报。' },
      { type: 'spacer', lines: 3 },
      { type: 'signature_block', authority: 'XX省人民政府', date: '2026年6月13日' },
      { type: 'spacer', lines: 1 },
      { type: 'attachment_note', attachments: ['2026年防汛重点区域清单', '防汛应急预案（修订版）'] },
      { type: 'recipient_line', text: '抄送：省委各部门，省人大常委会办公厅，省政协办公厅，省法院，省检察院。', recipientType: 'cc' },
      { type: 'centered_text', text: 'XX省人民政府办公厅        2026年6月13日印发', font_size_pt: 10, bold: false },
    ],
  };

  writeFileSync('official-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created official-content.json (公文示例)');
  printInitNext('official');
}

function printInitNext(type: DocumentType) {
  console.log('   Next steps:');
  console.log(`     1. Edit the generated JSON with your content`);
  console.log(`     2. thesis parse <your-template.docx> --type ${type}`);
  console.log(`     3. thesis build <content.json> -t <template.json>`);
}

// ── Help ──────────────────────────────────────────────────

function printHelp() {
  console.log(`
  OpenThesis — AI-powered Document Template Engine
  ─────────────────────────────────────────────────

  Parse any .docx template. Write in structured JSON.
  Output submission-ready DOCX. For thesis, journal
  articles, and official documents (公文).

  Commands:
    thesis parse <template.docx>     Parse template → JSON style DSL
    thesis build <content.json>       Render document → .docx
    thesis init [--type <t>]          Create sample content file

  Document types (--type):
    thesis   — 学位论文 (default)
    journal  — 期刊论文
    official — 公文 (GB/T 9704)

  Examples:
    thesis parse 清华博士模板.docx --type thesis --org "清华大学"
    thesis parse elsevier-template.docx --type journal
    thesis parse 公文模板.docx --type official --org "XX省人民政府"

    thesis build thesis-content.json -t template.json -o output.docx

  Options (build):
    -t <template.json>    Template file (default: built-in USTB)
    -o <output.docx>      Output path (default: input + .docx)
  `);
}

main().catch(e => { console.error(e); process.exit(1); });
