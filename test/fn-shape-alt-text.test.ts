// getShapeDescription / setShapeDescription / getShapeAltTitle /
// setShapeAltTitle — alt-text accessors for screen readers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeAltTitle,
  getShapeDescription,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAltTitle,
  setShapeDescription,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: alt-text accessors', () => {
  it('round-trips a description', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getShapeDescription(s)).toBeNull();
    setShapeDescription(s, 'A red rectangle marking the focus area.');
    expect(getShapeDescription(s)).toBe('A red rectangle marking the focus area.');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeDescription(shapes[shapes.length - 1]!)).toBe(
      'A red rectangle marking the focus area.',
    );
  });

  it('round-trips an alt title', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'hi',
    });
    setShapeAltTitle(s, 'Greeting');
    expect(getShapeAltTitle(s)).toBe('Greeting');
  });

  it('passing null clears the attribute', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeDescription(s, 'set');
    setShapeDescription(s, null);
    expect(getShapeDescription(s)).toBeNull();

    setShapeAltTitle(s, 'set');
    setShapeAltTitle(s, null);
    expect(getShapeAltTitle(s)).toBeNull();
  });
});
