// getShapeStrokeColor / getShapeStrokeWidth — sugar over getShapeStroke.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  emu,
  getShapeStrokeColor,
  getShapeStrokeWidth,
  getSlides,
  inches,
  loadPresentation,
  setShapeNoStroke,
  setShapeStroke,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeStrokeColor / getShapeStrokeWidth', () => {
  it('returns the color and width set via setShapeStroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeStroke(rect, { color: '#112233', widthEmu: emu(25400) });
    expect(getShapeStrokeColor(rect)).toBe('#112233');
    expect(getShapeStrokeWidth(rect)).toBe(25400);
  });

  it('returns null after clearing the stroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeNoStroke(rect);
    expect(getShapeStrokeColor(rect)).toBeNull();
    expect(getShapeStrokeWidth(rect)).toBeNull();
  });
});
