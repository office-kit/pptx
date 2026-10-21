import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import { builtinModules } from 'node:module';

const dist = new URL('../dist/', import.meta.url);
const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const SPECIFIER = /(?:\bfrom\s*|\bimport\s*\(?\s*)['"]([^'"./][^'"]*)['"]/g;

const packageName = (specifier) =>
  specifier
    .split('/')
    .slice(0, specifier.startsWith('@') ? 2 : 1)
    .join('/');

// The editor is bundled because it is not on npm; an import left in dist would make
// `npm install @office-kit/pptx-dev` resolve a package it does not declare.
test('dist imports only declared dependencies', async () => {
  const declared = new Set(Object.keys(manifest.dependencies));
  const files = (await readdir(dist, { recursive: true })).filter((name) =>
    /\.(mjs|d\.mts)$/.test(name),
  );
  assert.ok(files.length > 0);
  const undeclared = new Set();
  for (const file of files)
    for (const [, specifier] of (await readFile(new URL(file, dist), 'utf8')).matchAll(SPECIFIER)) {
      const name = packageName(specifier);
      if (!specifier.startsWith('node:') && !builtinModules.includes(name) && !declared.has(name))
        undeclared.add(`${file}: ${specifier}`);
    }
  assert.deepEqual([...undeclared], []);
});
