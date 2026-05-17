// Read-back for stroke dash + arrowhead.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideLine,
  addSlideShape,
  getShapeStrokeArrow,
  getShapeStrokeDash,
  getSlides,
  inches,
  loadPresentation,
  setShapeStroke,
  setShapeStrokeArrow,
  setShapeStrokeDash,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: stroke dash / arrow read-back', () => {
  it('getShapeStrokeDash returns the configured preset', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    expect(getShapeStrokeDash(shape)).toBeNull();
    setShapeStroke(shape, { color: '#000000' });
    setShapeStrokeDash(shape, 'dashDot');
    expect(getShapeStrokeDash(shape)).toBe('dashDot');
  });

  it('getShapeStrokeArrow returns null when no arrowhead is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(2), y: inches(2) },
    });
    expect(getShapeStrokeArrow(line, 'head')).toBeNull();
    expect(getShapeStrokeArrow(line, 'tail')).toBeNull();
  });

  it('round-trips an arrowhead', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(2), y: inches(2) },
    });
    setShapeStrokeArrow(line, 'tail', { type: 'triangle', width: 'med', length: 'lg' });
    const got = getShapeStrokeArrow(line, 'tail');
    expect(got).toEqual({ type: 'triangle', width: 'med', length: 'lg' });
  });
});
