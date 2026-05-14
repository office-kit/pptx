// getShapeStroke introspection.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearShapeStroke,
  getShapeStroke,
  getSlides,
  inches,
  loadPresentation,
  setShapeNoStroke,
  setShapeStroke,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeStroke', () => {
  it('returns inherit when spPr has no <a:ln>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    expect(getShapeStroke(shape).kind).toBe('inherit');
  });

  it('reads back a solid color outline + width', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeStroke(shape, { color: '#ABCDEF', widthEmu: 12700 });
    const s = getShapeStroke(shape);
    expect(s).toEqual({ kind: 'solid', color: '#ABCDEF', widthEmu: 12700 });
  });

  it('reports none for setShapeNoStroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeNoStroke(shape);
    expect(getShapeStroke(shape).kind).toBe('none');
  });

  it('returns inherit after clearShapeStroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeStroke(shape, { color: '#000000' });
    clearShapeStroke(shape);
    expect(getShapeStroke(shape).kind).toBe('inherit');
  });
});
