// getSlideLayoutNames — every layout name in part-name order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideLayoutNames,
  getSlideLayouts,
  getSlideLayoutName,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutNames', () => {
  it('matches getSlideLayouts(...).map(getSlideLayoutName)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const expected = getSlideLayouts(pres).map((l) => getSlideLayoutName(l));
    expect(getSlideLayoutNames(pres)).toEqual(expected);
  });

  it('includes the stock PowerPoint layout names', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const names = getSlideLayoutNames(pres);
    expect(names).toContain('Title and Content');
    expect(names).toContain('Blank');
  });
});
