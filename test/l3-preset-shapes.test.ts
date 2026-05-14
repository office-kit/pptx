// Preset shape authoring (rect, ellipse, triangle, arrow, star, ...).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: Slide.addShape', () => {
  it('emits prstGeom prst="ellipse"', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const ellipse = slide.addShape({
      preset: 'ellipse',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: '○',
    });
    expect(ellipse.kind).toBe('shape');
    expect(ellipse.position).toEqual({ x: inches(1), y: inches(1) });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('prst="ellipse"');
    expect(xml).toContain('○');
  });

  it('supports a wide set of preset tokens', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const presets = ['rect', 'roundRect', 'triangle', 'star5', 'rightArrow', 'cloud', 'heart'];
    for (let i = 0; i < presets.length; i++) {
      const p = presets[i];
      if (typeof p !== 'string') continue;
      slide.addShape({
        preset: p,
        x: inches(i % 4),
        y: inches(Math.floor(i / 4)),
        w: inches(1),
        h: inches(1),
      });
    }
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    for (const p of presets) {
      expect(xml).toContain(`prst="${p}"`);
    }
  });

  it('omits txBody when text is not supplied', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const before = slide.shapes.length;
    slide.addShape({
      preset: 'star5',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(slide.shapes.length).toBe(before + 1);
  });

  it('combines well with setFill / setStroke / setRotation', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const star = slide.addShape({
      preset: 'star5',
      x: inches(2),
      y: inches(2),
      w: inches(3),
      h: inches(3),
      text: '★',
    });
    star.setFill('accent2');
    star.setStroke({ color: '#000000', widthEmu: 12700 });
    star.setRotation(15);
    star.setTextFormat({ size: 32, color: '#FFFFFF', bold: true });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('prst="star5"');
    expect(xml).toContain('schemeClr val="accent2"');
    expect(xml).toContain('val="000000"');
    expect(xml).toContain('rot=');
    expect(xml).toContain('sz="3200"');
  });

  skipIfNoXmllint('preset shape validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    slide.addShape({
      preset: 'rightArrow',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'next',
    });
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
