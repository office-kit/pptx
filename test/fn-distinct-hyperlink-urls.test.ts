// getDistinctHyperlinkUrls — deduped URL list in first-seen order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getDistinctHyperlinkUrls,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getDistinctHyperlinkUrls', () => {
  it('reports each URL once in first-seen order', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    {
      const slide = addSlide(pres, { layout: blank });
      const a = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'A',
      });
      setShapeHyperlink(a, 'https://example.com/first');
    }
    {
      const slide = addSlide(pres, { layout: blank });
      const b = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'B',
      });
      setShapeHyperlink(b, 'https://example.com/second');
      const c = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'C',
      });
      setShapeHyperlink(c, 'https://example.com/first'); // duplicate
    }
    expect(getDistinctHyperlinkUrls(pres)).toEqual([
      'https://example.com/first',
      'https://example.com/second',
    ]);
  });

  it('returns empty when no link exists', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getDistinctHyperlinkUrls(pres)).toEqual([]);
  });
});
