// `getUnusedSlideMasters(pres)` — master part names no slide chains to.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideMasterPartNames,
  getSlideMasterUsageCounts,
  getUnusedSlideMasters,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getUnusedSlideMasters', () => {
  it('every returned name has a usage count of 0', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const unused = getUnusedSlideMasters(pres);
    const counts = getSlideMasterUsageCounts(pres);
    for (const name of unused) {
      expect(counts[name]).toBe(0);
    }
  });

  it('result is a subset of every master in the package', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const all = new Set(getSlideMasterPartNames(pres));
    for (const name of getUnusedSlideMasters(pres)) {
      expect(all.has(name)).toBe(true);
    }
  });
});
