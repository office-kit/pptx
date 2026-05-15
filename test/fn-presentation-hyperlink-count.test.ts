// getPresentationHyperlinkCount — fast counter for hyperlinks across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getAllHyperlinks,
  getPresentationHyperlinkCount,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationHyperlinkCount', () => {
  it('matches getAllHyperlinks length', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    {
      const slide = addSlide(pres, { layout: blank });
      const a = addSlideTextBox(slide, {
        x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'A',
      });
      setShapeHyperlink(a, 'https://example.com/a');
    }
    {
      const slide = addSlide(pres, { layout: blank });
      const b = addSlideTextBox(slide, {
        x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'B',
      });
      setShapeHyperlink(b, 'https://example.com/b');
    }
    expect(getPresentationHyperlinkCount(pres)).toBe(getAllHyperlinks(pres).length);
    expect(getPresentationHyperlinkCount(pres)).toBeGreaterThanOrEqual(2);
  });

  it('returns 0 on a deck with no hyperlinks', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getPresentationHyperlinkCount(pres)).toBe(0);
  });
});
