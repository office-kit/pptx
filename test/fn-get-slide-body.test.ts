// getSlideBody — read from the body placeholder.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideBody,
  loadPresentation,
  setSlideBody,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideBody', () => {
  it('returns the text written via setSlideBody', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlideBody(slide, 'one\ntwo');
    expect(getSlideBody(slide)).toBe('one\ntwo');
  });

  it('returns null on layouts without a body placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    expect(getSlideBody(slide)).toBeNull();
  });
});
