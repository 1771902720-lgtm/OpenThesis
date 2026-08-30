#!/usr/bin/env node

import { existsSync, readFileSync } from 'node:fs';
import { dirname, parse, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const minimumNodeMajor = 22;
const nodeMajor = Number(process.versions.node.split('.')[0]);
if (nodeMajor < minimumNodeMajor) {
  console.error(`OpenThesis requires Node.js ${minimumNodeMajor} or newer; found ${process.versions.node}.`);
  process.exit(1);
}

const explicitRoot = process.env.OPENTHESIS_ROOT;
const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const root = explicitRoot
  ? validateRoot(resolve(explicitRoot))
  : findRoot(process.cwd()) ?? findRoot(scriptDirectory);

if (!root) {
  console.error('Could not find an OpenThesis checkout. Run inside the repository or set OPENTHESIS_ROOT.');
  process.exit(1);
}

const cli = resolve(root, 'packages/cli/dist/index.js');
if (!existsSync(cli)) {
  console.error(`OpenThesis CLI is not built at ${cli}.`);
  console.error(`Run: cd ${root} && pnpm install --frozen-lockfile && pnpm build`);
  process.exit(1);
}

const result = spawnSync(process.execPath, [cli, ...process.argv.slice(2)], {
  cwd: process.cwd(),
  stdio: 'inherit',
});
if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}
process.exit(result.status ?? 1);

function findRoot(start) {
  let current = resolve(start);
  const filesystemRoot = parse(current).root;
  while (true) {
    const valid = validateRoot(current);
    if (valid) return valid;
    if (current === filesystemRoot) return undefined;
    current = dirname(current);
  }
}

function validateRoot(candidate) {
  const packagePath = resolve(candidate, 'package.json');
  if (!existsSync(packagePath)) return undefined;
  try {
    const packageJson = JSON.parse(readFileSync(packagePath, 'utf8'));
    return packageJson.name === 'openthesis' ? candidate : undefined;
  } catch {
    return undefined;
  }
}
