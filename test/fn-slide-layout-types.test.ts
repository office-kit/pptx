// getSlideLayoutTypes — every layout's <p:sldLayout type> token.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideLayoutTypes,
  getSlideLayouts,
  getSlideLayoutType,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutTypes', () => {
  it('matches getSlideLayouts(...).map(getSlideLayoutType)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const expected = getSlideLayouts(pres).map((l) => getSlideLayoutType(l));
    expect(getSlideLayoutTypes(pres)).toEqual(expected);
  });

  it('includes the stock spec tokens for the Office theme', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const types = getSlideLayoutTypes(pres);
    expect(types).toContain('title');
    expect(types).toContain('blank');
  });
});
