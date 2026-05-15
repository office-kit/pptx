// getSlideHyperlinkUrls — slide-level URL list in document order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideHyperlinkUrls,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideHyperlinkUrls', () => {
  it('returns URLs in document order, including duplicates', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const a = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'A',
    });
    setShapeHyperlink(a, 'https://example.com/x');
    const b = addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'B',
    });
    setShapeHyperlink(b, 'https://example.com/y');
    const c = addSlideTextBox(slide, {
      x: inches(0), y: inches(2), w: inches(2), h: inches(1), text: 'C',
    });
    setShapeHyperlink(c, 'https://example.com/x'); // duplicate
    expect(getSlideHyperlinkUrls(slide)).toEqual([
      'https://example.com/x',
      'https://example.com/y',
      'https://example.com/x',
    ]);
  });

  it('returns empty when no shape links out', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'no link',
    });
    expect(getSlideHyperlinkUrls(slide)).toEqual([]);
  });
});
