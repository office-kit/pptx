import { readFile } from 'node:fs/promises';
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
  setShapeShadow,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

describe('imported effect opacity transforms', () => {
  it.each([
    ['<a:alpha val="50%"/>', 0.5],
    ['<a:alphaMod val="50000"/>', 0.5],
    ['<a:alpha val="80000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/>', 0.5],
    ['<a:alphaOff val="-25000"/>', 0.75],
  ])('retains shadow opacity for %s through save/reload', async (transforms, opacity) => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
    );
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeShadow(shape, { color: '#FF0000', blurEmu: 0, offsetEmu: 0, angleDeg: 0 });
    const zip = readZip(await savePresentation(pres));
    const entry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
    const xml = new TextDecoder().decode(entry.data);
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
    const data = new TextEncoder().encode(
      xml.replace('<a:srgbClr val="FF0000"/>', `<a:srgbClr val="FF0000">${transforms}</a:srgbClr>`),
    );
    const imported = await loadPresentation(
      writeZip(
        zip.entries.map((candidate) => (candidate === entry ? { ...candidate, data } : candidate)),
      ),
    );
    const reloaded = await loadPresentation(await savePresentation(imported));
    for (const presentation of [imported, reloaded]) {
      const savedShape = getSlideShapes(getSlides(presentation)[0]!).at(-1)!;
      expect(getShapeEffects(presentation, savedShape)[0]).toMatchObject({
        color: '#FF0000',
        opacity,
      });
      expect(getShapeEffect(savedShape)).toMatchObject({ color: '#FF0000', opacity });
    }
  });
});
