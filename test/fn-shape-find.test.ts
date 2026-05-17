// Free-function shape lookup helpers.
//
// `findShapeByName`, `findShapesByName`, `findShapesByKind`, and
// `findShapeInPresentation` build on `getSlideShapes` to cover the
// common "find me the shape named X" pattern without forcing the
// caller to write `.find(...)` on every shape walk.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  findShapeByName,
  findShapeInPresentation,
  findShapesByKind,
  findShapesByName,
  getShapeName,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: shape lookup helpers', () => {
  it('findShapeByName returns the first matching shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'tagged',
      name: 'Logo',
    });
    expect(getShapeName(tb)).toBe('Logo');

    const hit = findShapeByName(slide, 'Logo');
    expect(hit).not.toBeNull();
    expect(getShapeName(hit!)).toBe('Logo');
    expect(findShapeByName(slide, 'nothing')).toBeNull();
  });

  it('findShapesByName returns every match', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'A',
      name: 'Marker',
    });
    addSlideTextBox(slide, {
      x: inches(2),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'B',
      name: 'Marker',
    });
    const all = findShapesByName(slide, 'Marker');
    expect(all).toHaveLength(2);
  });

  it('findShapesByKind filters by shape kind', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const shapes = findShapesByKind(slide, 'shape');
    // At least the rectangle we just added.
    expect(shapes.length).toBeGreaterThan(0);
    for (const s of shapes) {
      expect(s).toBeDefined();
    }
  });

  it('findShapeById returns the shape with that OOXML id', async () => {
    const { findShapeById, getShapeId } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'hi',
      name: 'Spot',
    });
    const id = getShapeId(tb);
    expect(findShapeById(slide, id)).not.toBeNull();
    expect(findShapeById(slide, 9999)).toBeNull();
  });

  it('findShapeInPresentation walks every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideTextBox(second!, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'X',
      name: 'GlobalTag',
    });
    expect(findShapeInPresentation(pres, 'GlobalTag')).not.toBeNull();
    expect(findShapeInPresentation(pres, 'NoSuch')).toBeNull();
    // Sanity: lint silence on `first`.
    void first;
  });
});
