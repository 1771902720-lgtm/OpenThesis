import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const cli = resolve('packages/cli/dist/index.js');

function run(args, cwd = process.cwd()) {
  return spawnSync(process.execPath, [cli, ...args], { cwd, encoding: 'utf8' });
}

test('help exits successfully', () => {
  const result = run(['--help']);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /OpenThesis/);
});

test('parse rejects non-DOCX input before it can be overwritten', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'template.txt');
  writeFileSync(input, 'keep me');
  const result = run(['parse', input]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /must be a \.docx file/);
  assert.equal(readFileSync(input, 'utf8'), 'keep me');
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
