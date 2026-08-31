import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { buildTestPdf } from './helpers/pdf.mjs';

const cli = resolve('packages/cli/dist/index.js');

function run(args, cwd = process.cwd()) {
  return spawnSync(process.execPath, [cli, ...args], { cwd, encoding: 'utf8' });
}

test('help exits successfully', () => {
  const result = run(['--help']);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /OpenThesis/);
});

test('parse rejects unsupported template input before it can be overwritten', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'template.txt');
  writeFileSync(input, 'keep me');
  const result = run(['parse', input]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /must be a \.docx or \.pdf file/);
  assert.equal(readFileSync(input, 'utf8'), 'keep me');
});

test('parse infers a template JSON file from a text-based PDF', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'template.pdf');
  const output = join(dir, 'template.template.json');
  writeFileSync(input, buildTestPdf());
  const result = run(['parse', input, '--org', 'PDF University']);
  assert.equal(result.status, 0, result.stderr);
  const template = JSON.parse(readFileSync(output, 'utf8'));
  assert.equal(template.meta.sourceFormat, 'pdf');
  assert.equal(template.meta.organization, 'PDF University');
  assert.ok(Object.keys(template.styles).length >= 2);

  const content = join(dir, 'content.json');
  const document = join(dir, 'rendered.docx');
  writeFileSync(content, JSON.stringify({
    type: 'thesis', meta: { title: 'PDF template test' }, cover: [],
    sections: [{ id: 'intro', type: 'chapter', title: 'Introduction', content: [
      { type: 'paragraph', text: 'Rendered with a PDF-derived template.' },
    ] }],
  }));
  const build = run(['build', content, '-t', output, '-o', document]);
  assert.equal(build.status, 0, build.stderr);
  assert.equal(readFileSync(document).subarray(0, 2).toString('ascii'), 'PK');
});

test('build appends .docx for an extensionless content file', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content');
  const content = {
    type: 'thesis', meta: { title: 'CLI test' }, cover: [],
    sections: [{ id: 'intro', type: 'chapter', title: 'Intro', content: [] }],
  };
  writeFileSync(input, JSON.stringify(content));
  const result = run(['build', input]);
  assert.equal(result.status, 0, result.stderr);
  assert.doesNotThrow(() => JSON.parse(readFileSync(input, 'utf8')));
  assert.equal(existsSync(`${input}.docx`), true);
});

test('missing option values produce a clear error', () => {
  const result = run(['build', 'content.json', '-o']);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Missing value for -o/);
});

test('import converts Markdown into structured JSON', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'paper.md');
  const output = join(dir, 'paper.json');
  writeFileSync(input, `---\ntype: journal\ntitle: CLI Markdown\nauthors: [Agent]\n---\n## Methods\nReproducible method.`);
  const result = run(['import', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(readFileSync(output, 'utf8'));
  assert.equal(document.type, 'journal');
  assert.equal(document.sections[0].title, 'Methods');
});

test('build renders Markdown directly to DOCX', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'thesis.md');
  const output = join(dir, 'thesis.docx');
  writeFileSync(input, `---\ntitle: Markdown Build\nauthor: Test Agent\n---\n## Introduction\nDirect build works.`);
  const result = run(['build', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(existsSync(output), true);
  assert.equal(readFileSync(output).subarray(0, 2).toString('ascii'), 'PK');
});
