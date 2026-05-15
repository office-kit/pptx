// getSlideOutline — per-slide { index, title, body } snapshot.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  getSlideOutline,
  loadPresentation,
  setShapeText,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideOutline', () => {
  it('captures title + body per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    setSlideTitle(slide, 'Roadmap');
    const body = findSlidePlaceholder(slide, 'body');
    if (body) setShapeText(body, 'Q1: discovery\nQ2: build');

    const outline = getSlideOutline(pres);
    expect(outline.length).toBe(1);
    expect(outline[0]!.index).toBe(0);
    expect(outline[0]!.title).toBe('Roadmap');
    expect(outline[0]!.body).toContain('Q1');
  });

  it('reports null for missing title / body', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout });

    const outline = getSlideOutline(pres);
    expect(outline[0]!.title).toBeNull();
    expect(outline[0]!.body).toBeNull();
  });

  it('walks every slide in document order', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'A');
    setSlideTitle(addSlide(pres, { layout }), 'B');
    setSlideTitle(addSlide(pres, { layout }), 'C');
    const outline = getSlideOutline(pres);
    expect(outline.map((e) => e.title)).toEqual(['A', 'B', 'C']);
    expect(outline.map((e) => e.index)).toEqual([0, 1, 2]);
  });
});
