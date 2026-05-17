// findSlidesByHyperlink — substring/regex match across outbound URLs.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  findSlidesByHyperlink,
  getSlideIndex,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesByHyperlink', () => {
  it('matches by substring', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    // Slide 0: docs link
    {
      const slide = addSlide(pres, { layout: blank });
      const t = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'docs',
      });
      setShapeHyperlink(t, 'https://docs.example.com/api');
    }
    const docsIdx = 0;
    // Slide 1: blog link
    {
      const slide = addSlide(pres, { layout: blank });
      const t = addSlideTextBox(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(2),
        h: inches(1),
        text: 'blog',
      });
      setShapeHyperlink(t, 'https://blog.example.com/post');
    }
    const blogIdx = 1;
    const matches = findSlidesByHyperlink(pres, 'docs.example.com');
    const indices = matches.map((s) => getSlideIndex(pres, s));
    expect(indices).toContain(docsIdx);
    expect(indices).not.toContain(blogIdx);
  });

  it('matches by RegExp', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const t = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'docs',
    });
    setShapeHyperlink(t, 'https://example.com/api');
    expect(findSlidesByHyperlink(pres, /^https:\/\/example\.com/).length).toBe(1);
    expect(findSlidesByHyperlink(pres, /^http:/).length).toBe(0);
  });
});
