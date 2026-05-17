// isShapePlaceholder — distinguish placeholder shapes from regular
// decorative geometry.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  findSlidePlaceholder,
  getSlideShapes,
  inches,
  isShapePlaceholder,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: isShapePlaceholder', () => {
  it('is true for title / body placeholders on a layout-derived slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const title = findSlidePlaceholder(slide, 'title')!;
    const body = findSlidePlaceholder(slide, 'body')!;
    expect(isShapePlaceholder(title)).toBe(true);
    expect(isShapePlaceholder(body)).toBe(true);
  });

  it('is false for a freshly-added decorative shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(isShapePlaceholder(rect)).toBe(false);
  });

  it('partitions a slide into placeholders + free shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const shapes = getSlideShapes(slide);
    const placeholders = shapes.filter((s) => isShapePlaceholder(s));
    const free = shapes.filter((s) => !isShapePlaceholder(s));
    expect(placeholders.length).toBeGreaterThan(0);
    expect(free.length).toBeGreaterThan(0);
    expect(placeholders.length + free.length).toBe(shapes.length);
  });
});
