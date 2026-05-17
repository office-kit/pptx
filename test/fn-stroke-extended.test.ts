// Extended stroke readers: dash style, cap, join, compound, arrow ends.
// The setters already exist; these readers complete the round-trip surface
// so renderers can faithfully reproduce a shape's outline.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideLine,
  addSlideShape,
  getShapeStrokeArrow,
  getShapeStrokeCap,
  getShapeStrokeCompound,
  getShapeStrokeDash,
  getShapeStrokeJoin,
  getSlides,
  inches,
  loadPresentation,
  setShapeStrokeArrow,
  setShapeStrokeDash,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: extended stroke readers', () => {
  it('round-trips dash via setShapeStrokeDash', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(2),
    });
    setShapeStrokeDash(shape, 'lgDashDot');
    expect(getShapeStrokeDash(shape)).toBe('lgDashDot');
  });

  it('returns null for cap / join / compound when no attribute set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(2),
    });
    expect(getShapeStrokeCap(shape)).toBeNull();
    expect(getShapeStrokeJoin(shape)).toBeNull();
    expect(getShapeStrokeCompound(shape)).toBeNull();
  });

  it('reads tail arrow set by setShapeStrokeArrow on a connector', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(3), y: inches(3) },
    });
    setShapeStrokeArrow(line, 'tail', { type: 'triangle', width: 'lg', length: 'med' });
    const arrow = getShapeStrokeArrow(line, 'tail');
    expect(arrow).not.toBeNull();
    expect(arrow!.type).toBe('triangle');
    expect(arrow!.width).toBe('lg');
    expect(arrow!.length).toBe('med');
  });
});
