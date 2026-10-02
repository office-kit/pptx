// `getShapeEffects` — returns every effect on the shape's `<a:effectLst>`
// in document order, not just the first one. Renderers need the full
// list because PowerPoint composes shadow + glow + softEdge into a
// single filter stack.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeEffect,
  getShapeEffects,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeEffects', () => {
  it('returns an empty array when no effects are set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    expect(getShapeEffects(pres, shape)).toEqual([]);
  });

  it('reads the outer shadow set by setShapeShadow with all numeric fields', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeShadow(shape, { color: '#000000', angleDeg: 45, opacity: 0.5 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('outerShdw');
    if (effects[0]!.kind === 'outerShdw') {
      expect(effects[0]!.color).toBe('#000000');
      expect(effects[0]!.angleDeg).toBeCloseTo(45);
      expect(effects[0]!.opacity).toBeCloseTo(0.5, 3);
      expect(effects[0]!.blurEmu).toBeGreaterThan(0);
    }
  });

  it('reads the glow set by setShapeGlow', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 63500 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('glow');
    if (effects[0]!.kind === 'glow') {
      expect(effects[0]!.color).toBe('#FF0000');
      expect(effects[0]!.radiusEmu).toBe(63500);
    }
  });

  it('reads scRGB and HSL effect colors through the public reader', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 100 });
    setShapeShadow(shape, { color: '#0000FF', angleDeg: 0, opacity: 1 });
    const zip = readZip(await savePresentation(pres));
    const slideEntry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
    let xml = new TextDecoder().decode(slideEntry.data);
    xml = xml.replace('<a:srgbClr val="FF0000"/>', '<a:scrgbClr r="50%" g="50%" b="50%"/>');
    xml = xml.replace(
      '<a:srgbClr val="0000FF"/>',
      '<a:hslClr hue="7200000" sat="100%" lum="50%"/>',
    );
    const imported = await loadPresentation(
      writeZip(
        zip.entries.map((entry) =>
          entry.name === slideEntry.name
            ? { ...entry, data: new TextEncoder().encode(xml) }
            : entry,
        ),
      ),
    );
    const importedShape = getSlideShapes(getSlides(imported)[0]!).find(
      (candidate) => getShapeEffects(imported, candidate).length > 0,
    )!;
    const effects = getShapeEffects(imported, importedShape);
    expect(effects).toMatchObject([
      { kind: 'glow', color: '#BCBCBC' },
      { kind: 'outerShdw', color: '#00FF00' },
    ]);
    expect(getShapeEffect(importedShape)).toMatchObject({ kind: 'shadow', color: '#00FF00' });

    const reloaded = await loadPresentation(await savePresentation(imported));
    const restoredShape = getSlideShapes(getSlides(reloaded)[0]!).find(
      (candidate) => getShapeEffects(reloaded, candidate).length > 0,
    )!;
    expect(getShapeEffects(reloaded, restoredShape)).toMatchObject([
      { kind: 'glow', color: '#BCBCBC' },
      { kind: 'outerShdw', color: '#00FF00' },
    ]);
    expect(getShapeEffect(restoredShape)).toMatchObject({ kind: 'shadow', color: '#00FF00' });
  });
});
