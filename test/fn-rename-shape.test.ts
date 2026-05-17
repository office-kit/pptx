// renameShape — change a shape's cNvPr@name.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  findShapeByName,
  getShapeName,
  getSlides,
  inches,
  loadPresentation,
  renameShape,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: renameShape', () => {
  it('round-trips the new name via findShapeByName', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'OldName',
    });
    expect(getShapeName(shape)).toBe('OldName');
    renameShape(shape, 'NewName');
    expect(getShapeName(shape)).toBe('NewName');
    expect(findShapeByName(slide, 'NewName')).not.toBeNull();
    expect(findShapeByName(slide, 'OldName')).toBeNull();
  });

  it('accepts the empty string', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    renameShape(shape, '');
    expect(getShapeName(shape)).toBe('');
  });
});
