// Preset shape authoring (rect, ellipse, triangle, arrow, star, ...).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getShapeKind,
  getShapePosition,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeFill,
  setShapeRotation,
  setShapeStroke,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: addSlideShape', () => {
  it('emits prstGeom prst="ellipse"', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = addSlide(pres, { layout });
    const ellipse = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: '○',
    });
    expect(getShapeKind(ellipse)).toBe('shape');
    expect(getShapePosition(ellipse)).toEqual({ x: inches(1), y: inches(1) });

    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('prst="ellipse"');
    expect(xml).toContain('○');
  });

  it('supports a wide set of preset tokens', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = addSlide(pres, { layout });
    const presets = ['rect', 'roundRect', 'triangle', 'star5', 'rightArrow', 'cloud', 'heart'];
    for (let i = 0; i < presets.length; i++) {
      const p = presets[i];
      if (typeof p !== 'string') continue;
      addSlideShape(slide, {
        preset: p as never,
        x: inches(i % 4),
        y: inches(Math.floor(i / 4)),
        w: inches(1),
        h: inches(1),
      });
    }
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    for (const p of presets) expect(xml).toContain(`prst="${p}"`);
  });

  it('omits txBody when text is not supplied', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = addSlide(pres, { layout });
    const before = getSlideShapes(slide).length;
    addSlideShape(slide, {
      preset: 'star5',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(getSlideShapes(getSlides(pres).at(-1)!).length).toBe(before + 1);
  });

  it('combines well with setShapeFill / setShapeStroke / setShapeRotation', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = addSlide(pres, { layout });
    const star = addSlideShape(slide, {
      preset: 'star5',
      x: inches(2),
      y: inches(2),
      w: inches(3),
      h: inches(3),
      text: '★',
    });
    setShapeFill(star, 'accent2');
    setShapeStroke(star, { color: '#000000', widthEmu: 12700 });
    setShapeRotation(star, 15);
    setShapeTextFormat(star, { size: 32, color: '#FFFFFF', bold: true });

    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('prst="star5"');
    expect(xml).toContain('schemeClr val="accent2"');
    expect(xml).toContain('val="000000"');
    expect(xml).toContain('rot=');
    expect(xml).toContain('sz="3200"');
  });

  skipIfNoXmllint('preset shape validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = addSlide(pres, { layout });
    addSlideShape(slide, {
      preset: 'rightArrow',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'next',
    });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
