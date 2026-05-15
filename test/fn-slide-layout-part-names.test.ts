// getSlideLayoutPartNames — every layout's URI in the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideLayoutPartName,
  getSlideLayoutPartNames,
  getSlideLayouts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutPartNames', () => {
  it('matches getSlideLayouts(...).map(getSlideLayoutPartName)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlideLayoutPartNames(pres)).toEqual(
      getSlideLayouts(pres).map((l) => getSlideLayoutPartName(l)),
    );
  });

  it('every entry is an absolute /ppt/slideLayouts/ path', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    for (const p of getSlideLayoutPartNames(pres)) {
      expect(p).toMatch(/^\/ppt\/slideLayouts\/slideLayout\d+\.xml$/);
    }
  });
});
