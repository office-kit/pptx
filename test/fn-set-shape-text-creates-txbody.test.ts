// A preset shape authored without `text` has no <p:txBody>. PowerPoint always
// gives an autoshape one so you can click in and type; setShapeText /
// appendShapeText must therefore create the body on demand rather than throw.
// Regression for the editor's "can't type into an inserted shape" bug.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  appendShapeText,
  getShapeText,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const addBareRect = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    // no `text` — so no <p:txBody> is authored
  });

describe('fn API: setShapeText creates a missing txBody', () => {
  it('an autoshape authored without text starts empty', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addBareRect(getSlides(pres)[0]!);
    expect(getShapeText(shape)).toBe('');
  });

  it('setShapeText populates a shape that had no text body', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addBareRect(getSlides(pres)[0]!);
    setShapeText(shape, 'こんにちは');
    expect(getShapeText(shape)).toBe('こんにちは');
  });

  it('appendShapeText works on a shape that had no text body', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addBareRect(getSlides(pres)[0]!);
    appendShapeText(shape, 'line 1');
    appendShapeText(shape, 'line 2');
    expect(getShapeText(shape)).toBe('line 1\nline 2');
  });

  it('multi-line text survives a save / load round-trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addBareRect(getSlides(pres)[0]!);
    setShapeText(shape, 'title\nsubtitle');

    const reopened = await loadPresentation(await savePresentation(pres));
    const shapes = getSlides(reopened)[0]!;
    const texts = (await import('../src/api/index.ts')).getSlideShapes(shapes).map(getShapeText);
    expect(texts).toContain('title\nsubtitle');
  });
});
