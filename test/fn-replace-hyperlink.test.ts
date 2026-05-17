// replaceHyperlink — bulk URL migration across the deck.

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
  replaceHyperlink,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: replaceHyperlink', () => {
  it('repoints every exact-match hyperlink and counts updates', async () => {
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
      setShapeHyperlink(a, 'https://old.example.com');
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
      setShapeHyperlink(b, 'https://old.example.com');
      const c = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'C',
      });
      setShapeHyperlink(c, 'https://other.example.com');
    }
    const n = replaceHyperlink(pres, 'https://old.example.com', 'https://new.example.com');
    expect(n).toBe(2);
    const distinct = getDistinctHyperlinkUrls(pres).slice().sort();
    expect(distinct).toEqual(['https://new.example.com', 'https://other.example.com']);
  });

  it('returns 0 when nothing matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const t = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'x',
    });
    setShapeHyperlink(t, 'https://other.example.com');
    expect(
      replaceHyperlink(pres, 'https://does-not-match.example.com', 'https://new.example.com'),
    ).toBe(0);
  });
});
