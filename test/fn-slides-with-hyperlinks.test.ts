// getSlidesWithHyperlinks — slides that link out.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlidesWithHyperlinks,
  getSlideIndex,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesWithHyperlinks', () => {
  it('returns only slides with at least one hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    // Linked slide
    {
      const slide = addSlide(pres, { layout: blank });
      const t = addSlideTextBox(slide, {
        x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'L',
      });
      setShapeHyperlink(t, 'https://example.com/a');
    }
    const linkedIdx = 0;
    // Plain slide
    {
      const slide = addSlide(pres, { layout: blank });
      addSlideTextBox(slide, {
        x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'P',
      });
    }
    const plainIdx = 1;
    const out = getSlidesWithHyperlinks(pres);
    const indices = out.map((s) => getSlideIndex(pres, s));
    expect(indices).toContain(linkedIdx);
    expect(indices).not.toContain(plainIdx);
  });

  it('returns empty when no slide has hyperlinks', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlidesWithHyperlinks(pres)).toEqual([]);
  });
});
