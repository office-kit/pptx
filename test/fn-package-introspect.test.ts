// Escape-hatch package introspection.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { listPackageParts, loadPresentation, readPackagePart } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: listPackageParts / readPackagePart', () => {
  it('lists every part with its content type + size', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const parts = listPackageParts(pres);
    expect(parts.length).toBeGreaterThan(0);
    expect(parts.some((p) => p.name === '/ppt/presentation.xml')).toBe(true);
    for (const p of parts) {
      expect(p.contentType.length).toBeGreaterThan(0);
      expect(p.byteLength).toBeGreaterThan(0);
    }
  });

  it('readPackagePart returns bytes for known parts and null otherwise', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const bytes = readPackagePart(pres, '/ppt/presentation.xml');
    expect(bytes).not.toBeNull();
    expect(bytes!.byteLength).toBeGreaterThan(0);
    expect(readPackagePart(pres, '/nope.xml')).toBeNull();
  });
});
