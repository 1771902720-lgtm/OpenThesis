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
import { resolve, dirname } from 'path';
import { parseTemplate } from '@openthesis/template-parser';
import { renderLegacy, createUSTBTemplate, renderDocument } from '@openthesis/docx-renderer';
import type { DocumentTemplate, DocumentType, LegacyDocumentJSON, ThesisDocument, JournalArticle, OfficialDocument } from '@openthesis/document-schema';

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
  const docType = validateDocType(typeIdx >= 0 ? args[typeIdx + 1] : 'thesis');
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
    const contentJson = JSON.parse(readFileSync(contentPath, 'utf-8'));
    const contentDir = dirname(contentPath);

    let template: DocumentTemplate;
    if (templatePath) {
      template = JSON.parse(readFileSync(templatePath, 'utf-8'));
    } else {
      console.log('  Using built-in USTB template...');
      template = createUSTBTemplate();
    }

    let docType = 'document';
    if (contentJson && typeof contentJson === 'object' && 'type' in contentJson) {
      console.log(`  Structured document detected: ${contentJson.type}`);
      docType = contentJson.type === 'thesis' ? 'thesis'
        : contentJson.type === 'journal' ? 'journal article'
        : 'official document';
        
      await renderDocument({
        outputPath,
        template,
        document: contentJson,
        contentDir,
      });
    } else {
      console.log('  Legacy document format detected, falling back to renderLegacy...');
      docType = contentJson.title?.includes('期刊') ? 'journal article'
        : contentJson.title?.includes('公文') || contentJson.title?.includes('通知') ? 'official document'
        : 'document';
        
      await renderLegacy(contentJson as LegacyDocumentJSON, template, outputPath);
    }

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
  const docType = validateDocType(typeIdx >= 0 ? args[typeIdx + 1] : 'thesis');

  switch (docType) {
    case 'thesis':   createThesisSample(); break;
    case 'journal':  createJournalSample(); break;
    case 'official': createOfficialSample(); break;
    default:         createThesisSample(); break;
  }
}

function createThesisSample() {
  const content: ThesisDocument = {
    type: 'thesis',
    meta: {
      title: '示例学位论文标题',
      degree: 'master',
      date: '2026年6月',
    },
    cover: [
      { type: 'spacer', lines: 6 },
      { type: 'centered_text', text: 'XX大学', font_size_pt: 26, bold: true },
      { type: 'centered_text', text: 'University Name (English)', font_size_pt: 16, bold: false },
      { type: 'spacer', lines: 2 },
      { type: 'centered_text', text: '示例硕士学位论文', font_size_pt: 22, bold: true },
      { type: 'spacer', lines: 4 },
      { type: 'centered_text', text: '学科专业：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '研究方向：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '日    期：2026年6月', font_size_pt: 14, bold: false },
      { type: 'page_break' },
    ],
    sections: [
      {
        id: 'intro',
        type: 'chapter',
        title: '第一章  绪论',
        content: [
          { type: 'heading2', text: '1.1  研究背景' },
          { type: 'paragraph', text: '正文示例——宋体小四号字，1.5倍行距，首行缩进2字符。西文使用Times New Roman。' },
          { type: 'heading2', text: '1.2  研究目的与意义' },
          { type: 'paragraph', text: '在此撰写研究目的...' },
        ],
      },
      {
        id: 'theory',
        type: 'chapter',
        title: '第二章  理论基础',
        content: [
          { type: 'heading2', text: '2.1  基本方程' },
          { type: 'paragraph', text: '控制方程为：' },
          { type: 'equation', latex: 'E = mc^2' },
        ],
      },
      {
        id: 'conclusion',
        type: 'chapter',
        title: '第三章  结论',
        content: [
          { type: 'paragraph', text: '研究成果总结...' },
        ],
      },
    ],
    backMatter: {
      references: [
        { id: '1', text: '作者一, 作者二. 论文标题. 期刊名称, 2026.' }
      ]
    }
  };

  writeFileSync('thesis-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created thesis-content.json (学位论文示例)');
  printInitNext('thesis');
}

function createJournalSample() {
  const content: JournalArticle = {
    type: 'journal',
    meta: {
      title: '示例期刊论文标题',
      authors: [
        { name: '作者一', affiliations: [{ institution: 'XX大学' }] },
        { name: '作者二', affiliations: [{ institution: 'XX研究所' }] },
        { name: '通讯作者', isCorresponding: true, affiliations: [{ institution: 'XX大学' }] }
      ],
      abstract: '本文研究了XXX问题，采用XXX方法，得出XXX结论。',
      keywords: ['关键词一', '关键词二', '关键词三']
    },
    sections: [
      {
        id: 'intro',
        type: 'introduction',
        title: '1  Introduction',
        content: [
          { type: 'paragraph', text: 'Introduction content here. This is a sample journal article body text.' },
          { type: 'heading2', text: '1.1  Background' },
          { type: 'paragraph', text: 'Background and literature review content...' },
        ]
      },
      {
        id: 'methods',
        type: 'methods',
        title: '2  Methods',
        content: [
          { type: 'paragraph', text: 'Methodology description...' },
        ]
      },
      {
        id: 'results',
        type: 'results',
        title: '3  Results and Discussion',
        content: [
          { type: 'paragraph', text: 'Results analysis...' },
          { type: 'table', caption: 'Table 1  Experimental results', headers: ['Parameter', 'Value', 'Error'], data: [['p1', '0.123', '±0.001'], ['p2', '0.456', '±0.002']] },
        ]
      },
      {
        id: 'conclusion',
        type: 'conclusion',
        title: '4  Conclusion',
        content: [
          { type: 'paragraph', text: 'Concluding remarks...' },
        ]
      }
    ],
    backMatter: {
      references: [
        { id: '1', text: 'Author A, Author B. Title of paper. Journal Name, 2026, 100(1): 1-10.' },
        { id: '2', text: 'Author C. Title of book. Publisher, 2025.' }
      ]
    }
  };

  writeFileSync('journal-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created journal-content.json (期刊论文示例)');
  printInitNext('journal');
}

function createOfficialSample() {
  const content: OfficialDocument = {
    type: 'official',
    meta: {
      title: '关于做好2026年防汛工作的通知',
      issuingAuthority: 'XX省人民政府',
      documentNumber: 'X政发〔2026〕1号',
      documentCategory: '通知',
      primaryRecipients: ['各市、州人民政府', '省政府各部门'],
      ccRecipients: ['省委各部门', '省人大常委会办公厅', '省政协办公厅'],
      date: '2026年6月13日',
    },
    body: [
      { type: 'recipient_line', text: '各市、州人民政府，省政府各部门：', recipientType: 'primary' },
      { type: 'paragraph', text: '为切实做好2026年防汛工作，保障人民群众生命财产安全，现将有关事项通知如下：' },
      { type: 'heading1', text: '一、提高思想认识，压实防汛责任' },
      { type: 'paragraph', text: '各级各部门要充分认识当前防汛形势的严峻性，坚决克服麻痹思想和侥幸心理。要严格落实以行政首长负责制为核心的各项防汛责任制，确保责任到人、措施到位。' },
      { type: 'heading1', text: '二、加强监测预警，做好应急准备' },
      { type: 'paragraph', text: '气象、水文部门要加强监测预报，及时发布预警信息。各地要修订完善防汛应急预案，充实抢险救援队伍，储备充足的防汛物资。' },
    ],
    attachments: [
      {
        order: 1,
        title: '2026年防汛重点区域清单',
        content: [
          { type: 'paragraph', text: '1. XX水库周边区域' },
          { type: 'paragraph', text: '2. XX河流中下游低洼地带' }
        ]
      }
    ]
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

const VALID_DOC_TYPES: DocumentType[] = ['thesis', 'journal', 'official'];

function validateDocType(value: string | undefined): DocumentType {
  if (!value || !VALID_DOC_TYPES.includes(value as DocumentType)) {
    console.error(`\n\u274c Invalid document type: "${value || ''}"`);
    console.error(`   Valid types: ${VALID_DOC_TYPES.join(', ')}`);
    process.exit(1);
  }
  return value as DocumentType;
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
