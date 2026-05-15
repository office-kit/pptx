// clearAllHyperlinks — sanitize every outbound URL across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  clearAllHyperlinks,
  findSlideLayout,
  getAllHyperlinks,
  getDistinctHyperlinkUrls,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllHyperlinks', () => {
  it('strips every hyperlink and reports the count', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const a = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'a',
    });
    setShapeHyperlink(a, 'https://example.com/a');
    const b = addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'b',
    });
    setShapeHyperlink(b, 'https://example.com/b');
    expect(getAllHyperlinks(pres).length).toBe(2);
    expect(clearAllHyperlinks(pres)).toBe(2);
    expect(getAllHyperlinks(pres).length).toBe(0);
    expect(getDistinctHyperlinkUrls(pres)).toEqual([]);
  });

  it('returns 0 when nothing to clear', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(clearAllHyperlinks(pres)).toBe(0);
  });
});
