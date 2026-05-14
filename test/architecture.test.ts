import { readFile, readdir } from 'node:fs/promises';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

// Enforces the internal-layer DAG documented in the foundation plan.
// Every cross-module import in src/ must satisfy `allowed[from]`. New modules
// MUST be added to `allowed` — an unknown module is a failure, not an exemption.

const SRC = fileURLToPath(new URL('../src/', import.meta.url));

type Module = 'api' | `internal/${string}`;

const allowed: Record<Module, ReadonlyArray<Module> | 'any-internal'> = {
  api: 'any-internal',
  'internal/xml': [],
  'internal/opc': ['internal/xml'],
  'internal/parts': ['internal/opc', 'internal/xml'],
  'internal/drawingml': ['internal/xml'],
  'internal/presentationml': ['internal/drawingml', 'internal/parts', 'internal/xml'],
  // chartml additionally needs `internal/opc` for the ZIP writer it uses to
  // emit the embedded xlsx workbook each chart wraps.
  'internal/chartml': ['internal/drawingml', 'internal/opc', 'internal/parts', 'internal/xml'],
  'internal/diagram': ['internal/drawingml', 'internal/parts', 'internal/xml'],
  'internal/io': [
    'internal/opc',
    'internal/parts',
    'internal/presentationml',
    'internal/chartml',
    'internal/diagram',
    'internal/xml',
  ],
  'internal/validator': 'any-internal',
  'internal/quirks': 'any-internal',
};

const walkTs = async (dir: string, acc: string[] = []): Promise<string[]> => {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      await walkTs(full, acc);
    } else if (entry.isFile() && entry.name.endsWith('.ts') && !entry.name.endsWith('.test.ts')) {
      acc.push(full);
    }
  }
  return acc;
};

const moduleOf = (absPath: string): Module | null => {
  const rel = relative(SRC, absPath).split(sep).join('/');
  // Top-level entries (index.ts, node.ts) are part of the public API surface.
  if (rel === 'index.ts' || rel === 'node.ts') return 'api';
  if (rel.startsWith('api/')) return 'api';
  const m = rel.match(/^internal\/([^/]+)\//);
  if (m) return `internal/${m[1]}` as Module;
  return null;
};

const resolveImport = (fromFile: string, spec: string): string | null => {
  if (!spec.startsWith('.')) return null;
  return resolve(dirname(fromFile), spec);
};

// Matches `import ... from '...'` and `export ... from '...'`, including
// multi-line and `import type` forms. Side-effect-only and dynamic imports are
// not handled — we don't use those, and the test should fail loudly if we do.
const importRegex = /^\s*(?:import|export)\s[\s\S]*?from\s*['"]([^'"]+)['"]/gm;

describe('architecture: layer DAG', () => {
  it('no source file imports outside its allowed targets', async () => {
    const files = await walkTs(SRC);
    const violations: string[] = [];

    for (const file of files) {
      const fileMod = moduleOf(file);
      if (!fileMod) {
        violations.push(`${relative(SRC, file)} is in an unknown module`);
        continue;
      }
      const allow = allowed[fileMod];
      if (allow === undefined) {
        violations.push(
          `${relative(SRC, file)} :: module ${fileMod} is not registered in allowed[]`,
        );
        continue;
      }
      const src = await readFile(file, 'utf8');
      for (const m of src.matchAll(importRegex)) {
        const spec = m[1];
        if (typeof spec !== 'string') continue;
        const target = resolveImport(file, spec);
        if (target === null) continue;
        const targetMod = moduleOf(target);
        if (!targetMod) continue;
        if (targetMod === fileMod) continue;
        if (allow === 'any-internal') continue;
        if (!allow.includes(targetMod)) {
          violations.push(
            `${relative(SRC, file)}: imports ${targetMod} (allowed for ${fileMod}: [${allow.join(', ')}])`,
          );
        }
      }
    }

    expect(violations).toEqual([]);
  });

  it('every declared module has at least one source file', async () => {
    const files = await walkTs(SRC);
    const seen = new Set<Module>();
    for (const f of files) {
      const m = moduleOf(f);
      if (m) seen.add(m);
    }
    const declared = Object.keys(allowed) as Module[];
    const missing = declared.filter((m) => !seen.has(m));
    expect(missing).toEqual([]);
  });
});
