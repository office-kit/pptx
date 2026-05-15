// getSlideMasterCount — number of slide masters in the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { getSlideMasterCount, loadPresentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideMasterCount', () => {
  it('reports at least one master on real fixtures', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideMasterCount(pres)).toBeGreaterThanOrEqual(1);
  });

  it('matches across fixtures', async () => {
    const a = await loadPresentation(await readFile(fixture('blank.pptx')));
    const b = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // Both stock fixtures use one master.
    expect(getSlideMasterCount(a)).toBe(1);
    expect(getSlideMasterCount(b)).toBe(1);
  });
});
