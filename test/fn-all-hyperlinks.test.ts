// getAllHyperlinks — every external URL across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getAllHyperlinks,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllHyperlinks', () => {
  it('reports each linked shape exactly once, with slide index', async () => {
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
      setShapeHyperlink(a, 'https://example.com/a');
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
      setShapeHyperlink(b, 'https://example.com/b');
    }
    const out = getAllHyperlinks(pres);
    expect(out.length).toBe(2);
    const urls = out.map((e) => e.url).sort();
    expect(urls).toEqual(['https://example.com/a', 'https://example.com/b']);
    // slideIndex values are non-negative
    for (const e of out) expect(e.slideIndex).toBeGreaterThanOrEqual(0);
  });

  it('returns empty when no hyperlink exists', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getAllHyperlinks(pres)).toEqual([]);
  });
});
