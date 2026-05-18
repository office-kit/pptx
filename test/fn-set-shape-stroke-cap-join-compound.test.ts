// Writers for the three stroke attributes that already had readers but
// no setters: line cap (`<a:ln cap/>`), line join (the round/bevel/miter
// child variants), and compound line (`<a:ln cmpd/>`).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeStrokeCap,
  getShapeStrokeCompound,
  getShapeStrokeJoin,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeStroke,
  setShapeStrokeCap,
  setShapeStrokeCompound,
  setShapeStrokeJoin,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const addShape = async () => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const slide = getSlides(pres)[0]!;
  const sh = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(2),
    h: inches(1),
  });
  setShapeStroke(sh, { color: '#000000', widthEmu: 9525 });
  return { pres, sh };
};

describe('fn API: setShapeStrokeCap', () => {
  it('round-trips every cap variant', async () => {
    const { sh } = await addShape();
    for (const cap of ['rnd', 'sq', 'flat'] as const) {
      setShapeStrokeCap(sh, cap);
      expect(getShapeStrokeCap(sh)).toBe(cap);
    }
  });

  it('survives save/reload', async () => {
    const { pres, sh } = await addShape();
    setShapeStrokeCap(sh, 'rnd');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeStrokeCap(reShapes[reShapes.length - 1]!)).toBe('rnd');
  });

  it('clears the attribute when set to null', async () => {
    const { sh } = await addShape();
    setShapeStrokeCap(sh, 'sq');
    expect(getShapeStrokeCap(sh)).toBe('sq');
    setShapeStrokeCap(sh, null);
    expect(getShapeStrokeCap(sh)).toBeNull();
  });
});

describe('fn API: setShapeStrokeJoin', () => {
  it('round-trips round / bevel / miter through save/reload', async () => {
    const { pres, sh } = await addShape();
    setShapeStrokeJoin(sh, 'miter');
    expect(getShapeStrokeJoin(sh)).toBe('miter');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeStrokeJoin(reShapes[reShapes.length - 1]!)).toBe('miter');
  });

  it('replaces the prior join child each call', async () => {
    const { sh } = await addShape();
    setShapeStrokeJoin(sh, 'round');
    setShapeStrokeJoin(sh, 'bevel');
    expect(getShapeStrokeJoin(sh)).toBe('bevel');
    setShapeStrokeJoin(sh, null);
    expect(getShapeStrokeJoin(sh)).toBeNull();
  });
});

describe('fn API: setShapeStrokeCompound', () => {
  it('round-trips every compound variant', async () => {
    const { sh } = await addShape();
    for (const cmpd of ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'] as const) {
      setShapeStrokeCompound(sh, cmpd);
      expect(getShapeStrokeCompound(sh)).toBe(cmpd);
    }
  });

  it('survives save/reload', async () => {
    const { pres, sh } = await addShape();
    setShapeStrokeCompound(sh, 'thickThin');
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeStrokeCompound(reShapes[reShapes.length - 1]!)).toBe('thickThin');
  });

  it('clears the attribute when set to null', async () => {
    const { sh } = await addShape();
    setShapeStrokeCompound(sh, 'dbl');
    setShapeStrokeCompound(sh, null);
    expect(getShapeStrokeCompound(sh)).toBeNull();
  });
});
