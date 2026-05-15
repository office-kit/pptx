// findSlideLayoutByType — look up a layout by its <p:sldLayout type="...">
// token. Counterpart to `findSlideLayout(pres, name)` that's locale-stable.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlideLayoutByType,
  getSlideLayoutName,
  getSlideLayoutType,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlideLayoutByType', () => {
  it('finds the "title" layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayoutByType(pres, 'title');
    expect(layout).not.toBeNull();
    expect(getSlideLayoutType(layout!)).toBe('title');
  });

  it('finds the "blank" layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayoutByType(pres, 'blank');
    expect(layout).not.toBeNull();
    expect(getSlideLayoutType(layout!)).toBe('blank');
    expect(getSlideLayoutName(layout!)).toBe('Blank');
  });

  it('returns null for an unknown type token', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findSlideLayoutByType(pres, 'noSuchType')).toBeNull();
  });
});
