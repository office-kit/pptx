// `getSlideMasterUsageCounts(pres)` — master part name → slide count.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideMasterPartNames,
  getSlideMasterUsageCounts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideMasterUsageCounts', () => {
  it('every master in the package appears as a key', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const masters = getSlideMasterPartNames(pres);
    const counts = getSlideMasterUsageCounts(pres);
    for (const m of masters) {
      expect(counts[m]).toBeDefined();
    }
  });

  it('reports a non-zero count for the master that the slides reference', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const counts = getSlideMasterUsageCounts(pres);
    const total = Object.values(counts).reduce((a, n) => a + n, 0);
    expect(total).toBeGreaterThan(0);
  });
});
